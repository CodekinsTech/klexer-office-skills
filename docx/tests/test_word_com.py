"""Checks docx_tool.py against files that Microsoft Word itself made, using Word's COM object model.

Skipped automatically when Word (Word.Application) is not installed or pywin32 is missing.
Run from the docx folder with the venv that has pywin32:
    python -m pytest tests\\test_word_com.py -v

How the checks work:
- Fixtures are built by Word through COM (tracked insertions and deletions in body and header,
  comments and a reply, a table with placeholders split across runs, formatting and
  paragraph-mark revisions). Word saves them as .docx.
- Every output of docx_tool.py is opened in Word. A prompt shown by Word while opening (a repair
  dialog) or an exception from Word counts as a failure.
- Expected values come from Word's object model (Revisions, Comments, Tables, Paragraphs, the
  final and original revision views), never from the tool's own output.
- Word's options are changed only for the test run (UseLocalUserInfo, so revision and comment
  authors are the names the test sets) and restored before Word quits. Word is quit in a finally
  block, and its process is stopped by PID if it does not exit.
"""

import importlib.util
import json
import shutil
import subprocess
import sys
from contextlib import contextmanager
import threading
import time
import zipfile
from pathlib import Path

import pytest

pytest.importorskip("win32com.client")
pytest.importorskip("win32gui")

import pythoncom  # noqa: E402  (installed with pywin32)
import win32com.client  # noqa: E402
import win32con  # noqa: E402
import win32gui  # noqa: E402
import win32process  # noqa: E402
from docx import Document  # noqa: E402
from pypdf import PdfReader  # noqa: E402

TOOL = Path(__file__).resolve().parents[1] / "scripts" / "docx_tool.py"

# Word constants (documented in the Word object model)
WD_MAIN_TEXT_STORY = 1
WD_PRIMARY_HEADER_STORY = 7
WD_PRIMARY_FOOTER_STORY = 9
WD_REVISION_INSERT = 1
WD_REVISION_DELETE = 2
WD_REVISION_PROPERTY = 3
WD_COLLAPSE_END = 0
WD_CHARACTER = 1
WD_WITH_IN_TABLE = 12
WD_STATISTIC_PAGES = 2
WD_REVISIONS_VIEW_FINAL = 0
WD_REVISIONS_VIEW_ORIGINAL = 1
WD_ALERTS_NONE = 0
WD_ALERTS_ALL = -1
WD_FORMAT_DOCX = 16
WD_FORMAT_PDF = 17
WD_DO_NOT_SAVE = 0

ANA = "Ana Lopez"
BEN = "Ben Ortiz"
TEST_BOT = "Test Bot"


def _word_registered() -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except OSError:
        return False


pytestmark = pytest.mark.skipif(not _word_registered(), reason="Microsoft Word is not installed")


def run_tool(cwd: Path, *args) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(TOOL), *[str(a) for a in args]],
        cwd=str(cwd), capture_output=True, text=True, encoding="utf-8",
    )


# --------------------------------------------------------------- Word session

class WordSession:
    """One Word instance for the whole module. Dialogs are watched and closed so they never hang a test."""

    def __init__(self):
        self.app = None
        self.pid = None
        self.dialogs: list[str] = []
        self._stop = threading.Event()
        self._orig_local_user = None

    def __enter__(self):
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = WD_ALERTS_NONE
        self._orig_local_user = self.app.Options.UseLocalUserInfo
        self.app.Options.UseLocalUserInfo = True
        blank = self.app.Documents.Add()
        self.pid = win32process.GetWindowThreadProcessId(blank.ActiveWindow.Hwnd)[1]
        blank.Close(WD_DO_NOT_SAVE)
        threading.Thread(target=self._watch, daemon=True).start()
        return self

    def __exit__(self, *exc):
        try:
            if self.app is not None and self._orig_local_user is not None:
                self.app.Options.UseLocalUserInfo = self._orig_local_user
        except Exception:  # noqa: BLE001 - restoring the option must not hide the test result
            pass
        finally:
            self._stop.set()
            try:
                if self.app is not None:
                    self.app.Quit(WD_DO_NOT_SAVE)
            except Exception:  # noqa: BLE001
                pass
            self._stop_process()
        return False

    def _stop_process(self):
        if self.pid is None:
            return
        for _ in range(20):
            if not _pid_alive(self.pid):
                return
            time.sleep(0.2)
        subprocess.run(["taskkill", "/PID", str(self.pid), "/F"], capture_output=True)

    def _watch(self):
        while not self._stop.is_set():
            try:
                win32gui.EnumWindows(self._scan, None)
            except Exception:  # noqa: BLE001 - the watcher must keep running
                pass
            time.sleep(0.2)

    def _scan(self, hwnd, _arg):
        try:
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid != self.pid or not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) not in ("#32770", "NUIDialog"):
                return True
            texts = []
            win32gui.EnumChildWindows(hwnd, lambda c, _x: texts.append(win32gui.GetWindowText(c)) or True, None)
            self.dialogs.append(" | ".join([win32gui.GetWindowText(hwnd)] + [t for t in texts if t]))
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:  # noqa: BLE001
            pass
        return True

    def open_checked(self, path: Path, read_only: bool = True):
        """Open and fail if Word shows a prompt (repair) or raises while opening.

        Read-only unless the caller edits the copy in memory; nothing is saved by this class.
        """
        self.dialogs.clear()
        self.app.DisplayAlerts = WD_ALERTS_ALL
        try:
            doc = self.app.Documents.Open(str(path), ReadOnly=read_only, AddToRecentFiles=False,
                                          ConfirmConversions=False)
        except pythoncom.com_error as exc:
            raise AssertionError(f"Word could not open {path.name}: {exc}; prompts: {self.dialogs}") from None
        finally:
            self.app.DisplayAlerts = WD_ALERTS_NONE
        assert not self.dialogs, f"Word showed a prompt when opening {path.name}: {self.dialogs}"
        return doc


def _pid_alive(pid: int) -> bool:
    out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True)
    return str(pid) in out.stdout


# ------------------------------------------------------- Word object helpers

def _find(story, text: str):
    """Return the range of the first exact match of `text` in `story` (raises if absent)."""
    rng = story
    assert rng.Find.Execute(FindText=text, MatchCase=True, Forward=True, Wrap=0), f"not found: {text!r}"
    return rng


def _set_author(session: WordSession, name: str):
    session.app.UserName = name
    session.app.UserInitials = "".join(part[0] for part in name.split())


def _set_view(doc, final: bool):
    view = doc.ActiveWindow.View
    view.ShowRevisionsAndComments = False
    view.RevisionsView = WD_REVISIONS_VIEW_FINAL if final else WD_REVISIONS_VIEW_ORIGINAL


def _stories(doc):
    """The main story plus every header and footer part of every section (Word's StoryRanges
    chain does not reach headers and footers, so they are read through the sections)."""
    yield doc.StoryRanges(WD_MAIN_TEXT_STORY)
    for s in range(1, doc.Sections.Count + 1):
        section = doc.Sections.Item(s)
        for parts in (section.Headers, section.Footers):
            for idx in (1, 2, 3):  # primary, first page, even pages
                part = parts.Item(idx)
                if part.Exists:
                    yield part.Range


def word_revisions(doc) -> list[tuple[int, str, str]]:
    """(type, author, text) for every revision in every story, from Word's Revisions collections."""
    rows = []
    for story in _stories(doc):
        coll = story.Revisions
        for i in range(1, coll.Count + 1):
            rv = coll.Item(i)
            rows.append((int(rv.Type), rv.Author, rv.Range.Text))
    return rows


def word_comments(doc) -> list[tuple[str, str, str]]:
    """(author, comment text, anchor text) for every comment, replies included."""
    rows = []
    for i in range(1, doc.Comments.Count + 1):
        c = doc.Comments.Item(i)
        rows.append((c.Author, c.Range.Text, c.Scope.Text))
    return rows


def word_body_paragraphs(doc, final: bool = True) -> list[str]:
    """Body paragraphs outside tables, in the final or original view."""
    _set_view(doc, final)
    texts = []
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs.Item(i)
        if not p.Range.Information(WD_WITH_IN_TABLE):
            texts.append(p.Range.Text.rstrip("\r"))
    return texts


def word_table_cells(doc, final: bool = True) -> list[list[str]]:
    _set_view(doc, final)
    rows = []
    for t in range(1, doc.Tables.Count + 1):
        table = doc.Tables.Item(t)
        for r in range(1, table.Rows.Count + 1):
            rows.append([table.Cell(r, c).Range.Text.rstrip("\r\x07") for c in range(1, table.Columns.Count + 1)])
    return rows


def word_header_footer(doc, final: bool = True) -> tuple[str, str]:
    _set_view(doc, final)
    section = doc.Sections.Item(1)
    return (section.Headers.Item(1).Range.Text.rstrip("\r"), section.Footers.Item(1).Range.Text.rstrip("\r"))


def word_pages(doc) -> int:
    return int(doc.ComputeStatistics(WD_STATISTIC_PAGES))


# ---------------------------------------------------------- fixture builders

def _build_track(session: WordSession, path: Path):
    app = session.app
    doc = app.Documents.Add()
    try:
        doc.Content.Text = "The annual fee is 100 dollars. Payment is due monthly.\rSecond paragraph stays put."
        doc.Sections.Item(1).Headers.Item(1).Range.Text = "Draft"
        doc.TrackRevisions = True
        _set_author(session, ANA)
        r = _find(doc.Content, "100 dollars")
        r.Collapse(WD_COLLAPSE_END)
        r.InsertAfter(" (revised)")
        _set_author(session, BEN)
        _find(doc.Content, "annual ").Delete()
        _find(doc.Content, "monthly").Text = "quarterly"
        _set_author(session, ANA)
        _find(doc.Content, "stays put").Delete()
        _set_author(session, BEN)
        _find(doc.Sections.Item(1).Headers.Item(1).Range, "Draft").Text = "Final"
        doc.TrackRevisions = False
        doc.SaveAs2(str(path), FileFormat=WD_FORMAT_DOCX)
    finally:
        doc.Close(WD_DO_NOT_SAVE)


def _build_comments(session: WordSession, path: Path):
    app = session.app
    doc = app.Documents.Add()
    try:
        doc.Content.Text = "The annual fee is 100 dollars. Payment is due monthly.\rSecond paragraph stays put."
        end = doc.Content
        end.Collapse(WD_COLLAPSE_END)
        table = doc.Tables.Add(end, 2, 2)
        table.Cell(1, 1).Range.Text = "Item"
        table.Cell(1, 2).Range.Text = "Price"
        table.Cell(2, 1).Range.Text = "Widget"
        table.Cell(2, 2).Range.Text = "9 dollars"
        _find(doc.Content, "due").Font.Bold = True
        _set_author(session, ANA)
        first = doc.Comments.Add(_find(doc.Content, "annual fee"), "Confirm the fee with finance.")
        _set_author(session, BEN)
        first.Replies.Add(first.Scope, "Confirmed at 100 dollars.")
        doc.Comments.Add(_find(doc.Content, "due monthly"), "Is monthly right?")
        _set_author(session, ANA)
        doc.Comments.Add(_find(doc.Tables.Item(1).Cell(2, 2).Range, "9 dollars"), "Check the currency.")
        _set_author(session, BEN)
        doc.Comments.Add(_find(doc.Content, "stays put"), "Keep this?")
        doc.SaveAs2(str(path), FileFormat=WD_FORMAT_DOCX)
    finally:
        doc.Close(WD_DO_NOT_SAVE)


def _build_table(session: WordSession, path: Path):
    app = session.app
    doc = app.Documents.Add()
    try:
        doc.Content.Text = "Dear {{NAME}}, invoice total {{TOTAL}}.\rDue date: "
        end = doc.Content
        end.Collapse(WD_COLLAPSE_END)
        table = doc.Tables.Add(end, 3, 2)
        table.Cell(1, 1).Range.Text = "Client"
        table.Cell(1, 2).Range.Text = "Ref"
        table.Cell(2, 1).Range.Text = "{{CLIENT}}"
        table.Cell(2, 2).Range.Text = "{{REF}}"
        table.Cell(3, 1).Range.Text = "Acme"
        table.Cell(3, 2).Range.Text = "R-100"
        _find(doc.Content, "NAME").Font.Bold = True
        _find(doc.Tables.Item(1).Cell(2, 1).Range, "CLIENT").Font.Italic = True
        section = doc.Sections.Item(1)
        section.Headers.Item(1).Range.Text = "Quarterly {{TITLE}}"
        section.Footers.Item(1).Range.Text = "Footer {{FOOT}} and {{REF}}"
        _set_author(session, ANA)
        doc.TrackRevisions = True
        last = _find(doc.Content, "Due date: ")
        last.Collapse(WD_COLLAPSE_END)
        last.InsertAfter("{{DATE}}")
        doc.TrackRevisions = False
        doc.SaveAs2(str(path), FileFormat=WD_FORMAT_DOCX)
    finally:
        doc.Close(WD_DO_NOT_SAVE)


def _build_fmt(session: WordSession, path: Path):
    app = session.app
    doc = app.Documents.Add()
    try:
        doc.Content.Text = "Rates are listed below.\rFinal line."
        doc.TrackRevisions = True
        _set_author(session, ANA)
        _find(doc.Content, "Rates").Font.Bold = True
        first = doc.Paragraphs.Item(1).Range
        first.InsertParagraphAfter()
        new_para = doc.Paragraphs.Item(2).Range
        new_para.InsertBefore("New paragraph.")
        _set_author(session, BEN)
        _find(doc.Content, "Final ").Delete()
        doc.TrackRevisions = False
        doc.SaveAs2(str(path), FileFormat=WD_FORMAT_DOCX)
    finally:
        doc.Close(WD_DO_NOT_SAVE)


def _build_review(session: WordSession, path: Path):
    """Comments and tracked changes together: a comment on text that is later deleted, and a
    tracked insertion inside a commented range."""
    app = session.app
    doc = app.Documents.Add()
    try:
        doc.Content.Text = "Our quote is 100 dollars for the widget."
        _set_author(session, ANA)
        first = doc.Comments.Add(_find(doc.Content, "100 dollars"), "Confirm the price.")
        _set_author(session, BEN)
        first.Replies.Add(first.Scope, "Price checked.")
        doc.Comments.Add(_find(doc.Content, "for the widget"), "Which widget?")
        doc.TrackRevisions = True
        _set_author(session, BEN)
        _find(doc.Content, "100 dollars").Text = "90 euros"
        _set_author(session, ANA)
        r = _find(doc.Content, "widget")
        r.Collapse(1)  # wdCollapseStart
        r.InsertAfter("big ")
        doc.TrackRevisions = False
        doc.SaveAs2(str(path), FileFormat=WD_FORMAT_DOCX)
    finally:
        doc.Close(WD_DO_NOT_SAVE)


def _build_insert_comment(session: WordSession, path: Path):
    """A comment on text that exists only as a tracked insertion."""
    app = session.app
    doc = app.Documents.Add()
    try:
        doc.Content.Text = "Plain text here."
        doc.TrackRevisions = True
        _set_author(session, BEN)
        r = _find(doc.Content, "text")
        r.Collapse(1)  # wdCollapseStart
        r.InsertAfter("XYZ ")
        doc.TrackRevisions = False
        _set_author(session, ANA)
        doc.Comments.Add(_find(doc.Content, "XYZ "), "Comment on insert only")
        doc.SaveAs2(str(path), FileFormat=WD_FORMAT_DOCX)
    finally:
        doc.Close(WD_DO_NOT_SAVE)


FIXTURES = {
    "word_track.docx": _build_track,
    "word_comments.docx": _build_comments,
    "word_table.docx": _build_table,
    "word_fmt.docx": _build_fmt,
    "word_review.docx": _build_review,
    "word_insert_comment.docx": _build_insert_comment,
}


def build_fixtures(session: WordSession, folder: Path) -> dict[str, Path]:
    folder.mkdir(parents=True, exist_ok=True)
    made = {}
    for name, builder in FIXTURES.items():
        path = folder / name
        builder(session, path)
        made[name] = path
    return made


def word_made(path: Path) -> bool:
    """True when the package says Word wrote it (docProps/app.xml names Microsoft Office Word)."""
    with zipfile.ZipFile(path) as zf:
        app_xml = zf.read("docProps/app.xml").decode("utf-8", "replace")
    return "Microsoft Office Word" in app_xml or "Microsoft Office" in app_xml


# ---------------------------------------------------------------- module fixtures

@pytest.fixture(scope="module")
def word():
    with WordSession() as session:
        yield session


@pytest.fixture(scope="module")
def word_files(word, tmp_path_factory):
    return build_fixtures(word, tmp_path_factory.mktemp("word_made"))


@pytest.fixture
def copy_of(tmp_path):
    """Copy a fixture into this test's folder so the shared fixture is never changed."""
    def _copy(src: Path) -> Path:
        dst = tmp_path / src.name
        shutil.copy(src, dst)
        return dst
    return _copy


def test_fixtures_are_word_made(word_files):
    for name, path in word_files.items():
        assert word_made(path), f"{name} was not written by Word"


def test_fixture_revisions_as_word_reports_them(word, word_files):
    doc = word.open_checked(word_files["word_track.docx"])
    try:
        rows = word_revisions(doc)
    finally:
        doc.Close(WD_DO_NOT_SAVE)
    assert sorted(rows) == sorted([
        (WD_REVISION_DELETE, BEN, "annual "),
        (WD_REVISION_INSERT, ANA, " (revised)"),
        (WD_REVISION_INSERT, BEN, "quarterly"),
        (WD_REVISION_DELETE, BEN, "monthly"),
        (WD_REVISION_DELETE, ANA, " stays put"),  # Word's smart cut takes the space before it
        (WD_REVISION_DELETE, BEN, "Draft"),
        (WD_REVISION_INSERT, BEN, "Final"),
    ])


# ------------------------------------------------------------------ helpers for checks

@contextmanager
def word_doc(session: WordSession, path: Path):
    doc = session.open_checked(path)
    try:
        yield doc
    finally:
        doc.Close(WD_DO_NOT_SAVE)


def tool_json(cwd: Path, *args):
    result = run_tool(cwd, *args)
    assert result.returncode == 0, result.stderr
    return json.loads(result.stdout)


def word_plain_revisions(rows):
    """Word insert/delete revisions as (kind, author, text); the paragraph mark is removed from the text."""
    kinds = {WD_REVISION_INSERT: "insert", WD_REVISION_DELETE: "delete"}
    return sorted((kinds[t], a, text.rstrip("\r")) for t, a, text in rows if t in kinds)


# ------------------------------------------------------------------- read / summary

@pytest.mark.parametrize("name", list(FIXTURES))
def test_read_matches_word(word, word_files, copy_of, name):
    src = copy_of(word_files[name])
    data = tool_json(src.parent, "read", "--json", src)
    with word_doc(word, src) as doc:
        assert [p["text"] for p in data["paragraphs"]] == word_body_paragraphs(doc, True)
        flat = [cell for table in data["tables"] for row in table for cell in row]
        assert flat == [cell for row in word_table_cells(doc, True) for cell in row]
        header, footer = word_header_footer(doc, True)
        assert [h for h in data["headers"] if h] == ([header] if header else [])
        assert [f for f in data["footers"] if f] == ([footer] if footer else [])


@pytest.mark.parametrize("name", list(FIXTURES))
def test_summary_matches_word(word, word_files, copy_of, name):
    src = copy_of(word_files[name])
    data = tool_json(src.parent, "summary", "--json", src)
    with word_doc(word, src) as doc:
        assert data["paragraph_count"] == len(word_body_paragraphs(doc, True))
        sizes = [[doc.Tables.Item(i).Rows.Count, doc.Tables.Item(i).Columns.Count]
                 for i in range(1, doc.Tables.Count + 1)]
        assert data["table_sizes"] == sizes
        assert data["has_comments"] == (doc.Comments.Count > 0)
        assert data["has_tracked_changes"] == (len(word_revisions(doc)) > 0)


# -------------------------------------------------------------- tracked changes

@pytest.mark.parametrize("name", ["word_track.docx", "word_review.docx", "word_insert_comment.docx"])
def test_track_show_matches_word(word, word_files, copy_of, name):
    src = copy_of(word_files[name])
    rows = tool_json(src.parent, "track-show", "--json", src)["revisions"]
    with word_doc(word, src) as doc:
        word_rows = word_revisions(doc)
    ours = sorted(
        ({"insert": WD_REVISION_INSERT, "delete": WD_REVISION_DELETE}[r["kind"]], r["author"], r["text"])
        for r in rows
    )
    assert ours == sorted(word_rows)
    assert len({r["id"] for r in rows}) == len(rows)  # Word's revision ids are unique
    if name == "word_track.docx":
        assert len(rows) == 7


def test_track_show_word_fmt_counts(word, word_files, copy_of):
    """Word merges a new paragraph and its text into one revision; the tool lists the paragraph mark
    as its own unsupported row (spec section 1). Plain rows must still agree with Word."""
    src = copy_of(word_files["word_fmt.docx"])
    rows = tool_json(src.parent, "track-show", "--json", src)["revisions"]
    with word_doc(word, src) as doc:
        word_rows = word_revisions(doc)
    plain = sorted((r["kind"], r["author"], r["text"]) for r in rows if r["kind"] != "unsupported")
    assert plain == word_plain_revisions(word_rows)
    details = sorted(r["detail"] for r in rows if r["kind"] == "unsupported")
    assert details == ["formatting change", "paragraph mark revision"]
    assert [t for t, _a, _x in word_rows].count(WD_REVISION_PROPERTY) == 1


def test_accept_all_matches_word(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_track.docx"])
    out = tmp_path / "accepted.docx"
    with word_doc(word, src) as doc:
        final_before = word_body_paragraphs(doc, True)
        header_before = word_header_footer(doc, True)
    result = run_tool(tmp_path, "track-resolve", src, "--mode", "accept", "--all", "--out", out, "--json")
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["resolved"] == 7
    with word_doc(word, out) as doc:
        assert word_revisions(doc) == []
        assert word_body_paragraphs(doc, True) == final_before
        assert word_header_footer(doc, True) == header_before


def test_reject_all_matches_word(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_track.docx"])
    out = tmp_path / "rejected.docx"
    with word_doc(word, src) as doc:
        original_before = word_body_paragraphs(doc, False)
        header_original = word_header_footer(doc, False)[0]
    result = run_tool(tmp_path, "track-resolve", src, "--mode", "reject", "--all", "--out", out)
    assert result.returncode == 0, result.stderr
    with word_doc(word, out) as doc:
        assert word_revisions(doc) == []
        assert word_body_paragraphs(doc, True) == original_before
        assert word_header_footer(doc, True)[0] == header_original == "Draft"


def test_accept_one_id_matches_word(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_track.docx"])
    rows = tool_json(tmp_path, "track-show", "--json", src)["revisions"]
    target = next(r for r in rows if r["kind"] == "insert" and r["author"] == ANA)
    out = tmp_path / "one.docx"
    result = run_tool(tmp_path, "track-resolve", src, "--mode", "accept", "--id", str(target["id"]), "--out", out)
    assert result.returncode == 0, result.stderr
    with word_doc(word, src) as before, word_doc(word, out) as after:
        expected = sorted(r for r in word_revisions(before) if r[1:] != (target["author"], target["text"]))
        assert sorted(word_revisions(after)) == expected
        assert word_body_paragraphs(after, True) == word_body_paragraphs(before, True)


def test_formatting_revision_refused_on_word_file(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_fmt.docx"])
    out = tmp_path / "refused.docx"
    result = run_tool(tmp_path, "track-resolve", src, "--mode", "accept", "--all", "--out", out)
    assert result.returncode == 2, result.stdout + result.stderr
    assert "cannot be resolved" in result.stderr
    assert not out.exists()


def test_plain_delete_accepted_on_word_file(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_fmt.docx"])
    rows = tool_json(tmp_path, "track-show", "--json", src)["revisions"]
    target = next(r for r in rows if r["kind"] == "delete")
    out = tmp_path / "fmt_one.docx"
    result = run_tool(tmp_path, "track-resolve", src, "--mode", "accept", "--id", str(target["id"]), "--out", out)
    assert result.returncode == 0, result.stderr
    with word_doc(word, src) as before, word_doc(word, out) as after:
        assert word_body_paragraphs(after, True) == word_body_paragraphs(before, True)
        assert sorted(word_revisions(after)) == sorted(
            r for r in word_revisions(before) if r[2].rstrip("\r") != target["text"])


# ------------------------------------------------------------- replace / fill

def test_replace_on_word_file(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_track.docx"])
    pairs = {"dollars": "USD", "Second": "Third", "revised": "edited", "Final": "Done", "monthly": "daily"}
    pairs_file = tmp_path / "pairs.json"
    pairs_file.write_text(json.dumps(pairs), encoding="utf-8")
    out = tmp_path / "replaced.docx"
    result = run_tool(tmp_path, "replace", src, "--pairs", pairs_file, "--out", out, "--json")
    assert result.returncode == 0, result.stderr
    report = json.loads(result.stdout)
    assert report["replacements"] == 2
    assert report["per_pair"] == {"dollars": 1, "Second": 1, "revised": 0, "Final": 0, "monthly": 0}
    assert report["skipped_in_tracked_change"] == 2
    assert report["remaining_after"]["revised"] == 1
    with word_doc(word, src) as before, word_doc(word, out) as after:
        assert sorted(word_revisions(after)) == sorted(word_revisions(before))
        assert word_body_paragraphs(after, True) == [
            "The fee is 100 USD (revised). Payment is due quarterly.", "Third paragraph."]
        assert word_header_footer(after, True)[0] == "Final"


def test_replace_in_word_table_and_footer(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_table.docx"])
    first = tmp_path / "table_replaced.docx"
    result = run_tool(tmp_path, "replace", src, "--find", "Acme", "--replace", "Globex", "--out", first, "--json")
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["replacements"] == 1
    out = tmp_path / "footer_replaced.docx"
    second = run_tool(tmp_path, "replace", first, "--find", "Footer", "--replace", "Foot", "--out", out, "--json")
    assert second.returncode == 0, second.stderr
    assert json.loads(second.stdout)["replacements"] == 1
    with word_doc(word, src) as before, word_doc(word, out) as after:
        assert word_table_cells(after, True)[2] == ["Globex", "R-100"]
        assert word_table_cells(after, True)[0] == ["Client", "Ref"]
        assert word_header_footer(after, True)[1] == "Foot {{FOOT}} and {{REF}}"
        assert sorted(word_revisions(after)) == sorted(word_revisions(before))


def test_fill_word_table_missing_and_allowed(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_table.docx"])
    values = {"NAME": "Ana", "CLIENT": "Globex", "REF": "R-7", "TITLE": "Report", "FOOT": "F1", "DATE": "1 Nov"}
    values_file = tmp_path / "values.json"
    values_file.write_text(json.dumps(values), encoding="utf-8")
    out = tmp_path / "filled.docx"
    refused = run_tool(tmp_path, "fill", src, "--values", values_file, "--out", out)
    assert refused.returncode == 1
    assert "TOTAL" in refused.stderr and "DATE" in refused.stderr
    assert not out.exists()

    result = run_tool(tmp_path, "fill", src, "--values", values_file, "--out", out, "--allow-unfilled", "--json")
    assert result.returncode == 0, result.stderr
    report = json.loads(result.stdout)
    assert report["unfilled"] == ["TOTAL", "DATE"]
    assert report["skipped_in_tracked_change"] == 1
    with word_doc(word, src) as before, word_doc(word, out) as after:
        assert word_body_paragraphs(after, True) == [
            "Dear Ana, invoice total {{TOTAL}}.", "Due date: {{DATE}}", ""]
        assert word_table_cells(after, True)[1] == ["Globex", "R-7"]
        assert word_header_footer(after, True) == ("Quarterly Report", "Footer F1 and R-7")
        assert sorted(word_revisions(after)) == sorted(word_revisions(before))


# ----------------------------------------------------------------- comments

def test_comments_list_matches_word(word, word_files, copy_of):
    src = copy_of(word_files["word_comments.docx"])
    rows = tool_json(src.parent, "comments-list", "--json", src)["comments"]
    with word_doc(word, src) as doc:
        word_rows = word_comments(doc)
    assert [(r["author"], r["text"], r["anchor_text"]) for r in rows] == word_rows
    assert len(rows) == 5


def test_comments_add_on_word_file(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_comments.docx"])
    out = tmp_path / "commented.docx"
    result = run_tool(tmp_path, "comments-add", src, "--anchor", "Payment is due", "--text", "Test comment",
                      "--author", TEST_BOT, "--out", out, "--json")
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["anchor_text"] == "Payment is due"
    with word_doc(word, src) as before, word_doc(word, out) as after:
        # Word lists comments in document order, so compare as sorted lists
        assert sorted(word_comments(after)) == sorted(word_comments(before) + [(TEST_BOT, "Test comment", "Payment is due")])
        assert word_body_paragraphs(after, True) == word_body_paragraphs(before, True)


def test_comments_add_ambiguous_then_occurrence(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_comments.docx"])
    ambiguous = run_tool(tmp_path, "comments-add", src, "--anchor", "dollars", "--text", "Rate?",
                         "--author", TEST_BOT, "--out", tmp_path / "x.docx")
    assert ambiguous.returncode == 1
    assert "occurs 2 times" in ambiguous.stderr
    out = tmp_path / "second.docx"
    result = run_tool(tmp_path, "comments-add", src, "--anchor", "dollars", "--text", "Rate?",
                      "--author", TEST_BOT, "--occurrence", "2", "--out", out)
    assert result.returncode == 0, result.stderr
    with word_doc(word, out) as after:
        assert len(word_comments(after)) == 6
        assert (TEST_BOT, "Rate?", "dollars") in word_comments(after)
        bot = next(c for c in range(1, after.Comments.Count + 1) if after.Comments.Item(c).Author == TEST_BOT)
        assert after.Comments.Item(bot).Scope.Information(WD_WITH_IN_TABLE), "comment should sit in the table cell"


# ------------------------------------------------------------- validate / control

@pytest.mark.parametrize("name", list(FIXTURES))
def test_validate_word_files(copy_of, word_files, tmp_path, name):
    src = copy_of(word_files[name])
    result = run_tool(tmp_path, "validate", "--json", src)
    assert result.returncode == 0, result.stdout
    assert json.loads(result.stdout)["verdict"] == "valid"


def test_word_detector_flags_truncated_file(word, word_files, copy_of):
    """Control: a damaged copy must fail the open check, so a passing check means something."""
    src = copy_of(word_files["word_track.docx"])
    data = src.read_bytes()
    src.write_bytes(data[: len(data) // 2])
    with pytest.raises(AssertionError):
        word.open_checked(src)


# ------------------------------------------- Word's own Accept All / Reject All

def word_resolve_in_memory(session: WordSession, path: Path, accept: bool):
    """Word's own Accept All or Reject All on an in-memory copy (nothing is saved)."""
    doc = session.open_checked(path, read_only=False)
    try:
        for story in _stories(doc):
            if accept:
                story.Revisions.AcceptAll()
            else:
                story.Revisions.RejectAll()
        return word_comments(doc), word_body_paragraphs(doc, True), word_header_footer(doc, True)
    finally:
        doc.Close(WD_DO_NOT_SAVE)


def assert_resolved_like_word(session, out: Path, comments_list, expected):
    expected_comments, expected_body, expected_header_footer = expected
    with word_doc(session, out) as doc:
        assert word_revisions(doc) == []
        assert word_comments(doc) == expected_comments
        assert word_body_paragraphs(doc, True) == expected_body
        assert word_header_footer(doc, True) == expected_header_footer
    assert [(c["author"], c["text"], c["anchor_text"]) for c in comments_list] == expected_comments


@pytest.mark.parametrize("name", ["word_track.docx", "word_review.docx", "word_insert_comment.docx"])
@pytest.mark.parametrize("mode", ["accept", "reject"])
def test_resolve_all_matches_word_own(word, word_files, copy_of, tmp_path, name, mode):
    src = copy_of(word_files[name])
    expected = word_resolve_in_memory(word, src, accept=(mode == "accept"))
    out = tmp_path / f"{mode}.docx"
    result = run_tool(tmp_path, "track-resolve", src, "--mode", mode, "--all", "--out", out)
    assert result.returncode == 0, result.stderr
    comments_list = tool_json(tmp_path, "comments-list", "--json", out)["comments"]
    assert_resolved_like_word(word, out, comments_list, expected)


@pytest.mark.parametrize("mode", ["accept", "reject"])
def test_comment_added_across_insertion_matches_word_own(word, word_files, copy_of, tmp_path, mode):
    """comments-add over tracked text, then resolve: the comment must survive or go the way Word's does."""
    src = copy_of(word_files["word_track.docx"])
    added = tmp_path / "added.docx"
    result = run_tool(tmp_path, "comments-add", src, "--anchor", "due quarterly", "--text", "Mixed anchor",
                      "--author", TEST_BOT, "--out", added)
    assert result.returncode == 0, result.stderr
    expected = word_resolve_in_memory(word, added, accept=(mode == "accept"))
    out = tmp_path / f"{mode}_added.docx"
    resolved = run_tool(tmp_path, "track-resolve", added, "--mode", mode, "--all", "--out", out)
    assert resolved.returncode == 0, resolved.stderr
    comments_list = tool_json(tmp_path, "comments-list", "--json", out)["comments"]
    assert_resolved_like_word(word, out, comments_list, expected)


# ----------------------------------------------------------------- to-pdf through Word

def word_process_ids() -> set[int]:
    listing = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, check=False)
    pids = set()
    for line in listing.stdout.splitlines():
        fields = line.strip().strip('"').split('","')
        if len(fields) > 1 and fields[1].isdigit():
            pids.add(int(fields[1]))
    return pids


@pytest.mark.parametrize("name", ["word_track.docx", "word_table.docx", "word_review.docx"])
def test_to_pdf_pages_match_word(word, word_files, copy_of, tmp_path, name):
    src = copy_of(word_files[name])
    with word_doc(word, src) as doc:
        expected_pages = word_pages(doc)
    out = tmp_path / "out.pdf"
    result = run_tool(tmp_path, "to-pdf", src, "--out", out)
    assert result.returncode == 0, result.stderr
    assert "converted with Microsoft Word" in result.stdout
    assert len(PdfReader(str(out)).pages) == expected_pages
    assert src.read_bytes() == word_files[name].read_bytes(), "the input must not change"
    assert not [p for p in tmp_path.iterdir() if p.name.startswith(".docx-tool")], "temp folder left behind"


@pytest.mark.parametrize("name", ["word_track.docx", "word_table.docx"])
def test_to_pdf_text_matches_word(word, word_files, copy_of, tmp_path, name):
    """Word cannot re-open a PDF through COM here (its PDF import waits on a prompt that COM never
    returns from), so the PDF is checked for its text. The PDF shows tracked changes as markup, so
    the first words come from Word's original view, which the markup keeps contiguous."""
    src = copy_of(word_files[name])
    with word_doc(word, src) as doc:
        first_words = " ".join(word_body_paragraphs(doc, False)[0].split()[:2])
    out = tmp_path / "text.pdf"
    result = run_tool(tmp_path, "to-pdf", src, "--out", out)
    assert result.returncode == 0, result.stderr
    text = " ".join(" ".join(page.extract_text() for page in PdfReader(str(out)).pages).split())
    assert first_words in text


def test_to_pdf_refused_without_word_or_libreoffice(tmp_path, monkeypatch, capsys):
    spec = importlib.util.spec_from_file_location("docx_tool_under_test", TOOL)
    tool = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = tool  # dataclasses in the tool look the module up by name
    spec.loader.exec_module(tool)
    monkeypatch.setattr(tool, "_word_available", lambda: False)
    monkeypatch.setattr(tool, "_find_soffice", lambda: None)
    src = tmp_path / "plain.docx"
    Document().save(str(src))
    code = tool.main(["to-pdf", str(src), "--out", str(tmp_path / "plain.pdf")])
    assert code == 2
    assert "LibreOffice is not installed. Install it or convert the file in Word." in capsys.readouterr().err
    assert not (tmp_path / "plain.pdf").exists()


def test_to_pdf_timeout_leaves_no_word_running(word, word_files, copy_of, tmp_path):
    src = copy_of(word_files["word_track.docx"])
    before = word_process_ids()
    result = run_tool(tmp_path, "to-pdf", src, "--out", tmp_path / "timeout.pdf", "--timeout", "1")
    assert result.returncode == 1
    assert "timed out after 1 seconds" in result.stderr
    assert not (tmp_path / "timeout.pdf").exists()
    deadline = time.time() + 10
    while word_process_ids() - before and time.time() < deadline:
        time.sleep(0.5)
    assert word_process_ids() - before == set(), "a Word instance started by to-pdf kept running"
