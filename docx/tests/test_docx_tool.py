"""Acceptance tests for docx_tool.py, one test per acceptance test in specs/docx.md section 3.

Every test builds its own input files with python-docx. Run from the docx folder with:
    python -m pytest tests -v
"""

import importlib.util
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from pypdf import PdfReader

TOOL = Path(__file__).resolve().parents[1] / "scripts" / "docx_tool.py"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def run_tool(cwd: Path, *args) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(TOOL), *[str(a) for a in args]],
        cwd=str(cwd), capture_output=True, text=True, encoding="utf-8",
    )


def read_json(cwd: Path, path) -> dict:
    result = run_tool(cwd, "read", "--json", path)
    assert result.returncode == 0, result.stderr
    return json.loads(result.stdout)


def _revision(tag, rid, author, text, deleted=False):
    wrapper = OxmlElement(f"w:{tag}")
    wrapper.set(qn("w:id"), str(rid))
    wrapper.set(qn("w:author"), author)
    wrapper.set(qn("w:date"), "2026-10-01T09:00:00Z")
    run = OxmlElement("w:r")
    node = OxmlElement("w:delText" if deleted else "w:t")
    node.text = text
    node.set(qn("xml:space"), "preserve")
    run.append(node)
    wrapper.append(run)
    return wrapper


def make_summary_doc(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("Spec", style="Heading 1")
    doc.add_paragraph("Scope", style="Heading 2")
    doc.add_paragraph("Para one")
    doc.add_paragraph("Para two")
    doc.add_table(rows=2, cols=2)
    doc.save(path)
    return path


def make_read_doc(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("Alpha")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "a"
    table.cell(0, 1).text = "b"
    table.cell(1, 0).text = "c"
    table.cell(1, 1).text = "d"
    doc.add_paragraph("Beta")
    doc.save(path)
    return path


def make_replace_doc(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("ACME signed")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "ACME"
    section = doc.sections[0]
    section.header.is_linked_to_previous = False
    section.header.paragraphs[0].text = "ACME"
    doc.save(path)
    return path


def make_fill_template(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("Dear {{NAME}}, due {{DATE}}.")
    doc.save(path)
    return path


def make_track_doc(path: Path) -> Path:
    doc = Document()
    para = doc.add_paragraph()
    para._p.append(_revision("ins", 1, "Ana", "new"))
    para._p.append(_revision("del", 2, "Ben", "old", deleted=True))
    doc.save(path)
    return path


def test_summary(tmp_path):
    make_summary_doc(tmp_path / "t_sum.docx")
    result = run_tool(tmp_path, "summary", "--json", "t_sum.docx")
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["headings"] == [{"level": 1, "text": "Spec"}, {"level": 2, "text": "Scope"}]
    assert data["paragraph_count"] == 4
    assert data["table_count"] == 1
    assert data["table_sizes"] == [[2, 2]]
    assert data["section_count"] == 1
    assert data["has_comments"] is False


def test_read(tmp_path):
    make_read_doc(tmp_path / "t_read.docx")
    result = run_tool(tmp_path, "read", "t_read.docx")
    assert result.returncode == 0, result.stderr
    assert result.stdout.splitlines() == ["Alpha", "a\tb", "c\td", "", "Beta"]


def test_create(tmp_path):
    (tmp_path / "t_outline.txt").write_text("# Title\n\nHello world\n- one\n- two\n", encoding="utf-8")
    result = run_tool(tmp_path, "create", "t_outline.txt", "--out", "t_out.docx")
    assert result.returncode == 0, result.stderr
    data = read_json(tmp_path, "t_out.docx")
    paragraphs = data["paragraphs"]
    assert [(p["style"], p["text"]) for p in paragraphs] == [
        ("Heading 1", "Title"),
        ("Normal", "Hello world"),
        ("List Bullet", "one"),
        ("List Bullet", "two"),
    ]


def test_replace(tmp_path):
    make_replace_doc(tmp_path / "t_rep.docx")
    result = run_tool(tmp_path, "replace", "t_rep.docx", "--find", "ACME", "--replace", "Globex", "--json")
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["replacements"] == 3
    out = tmp_path / "t_rep.replace.docx"
    assert out.exists()
    # read --json covers body paragraphs, tables and headers in one object
    text = json.dumps(read_json(tmp_path, out), ensure_ascii=False)
    assert text.count("ACME") == 0
    assert text.count("Globex") == 3


def test_fill(tmp_path):
    make_fill_template(tmp_path / "t_fill.docx")
    (tmp_path / "t_fill.json").write_text(json.dumps({"NAME": "Ana", "DATE": "1 Nov"}), encoding="utf-8")
    result = run_tool(tmp_path, "fill", "t_fill.docx", "--values", "t_fill.json")
    assert result.returncode == 0, result.stderr
    data = read_json(tmp_path, "t_fill.fill.docx")
    assert data["paragraphs"][0]["text"] == "Dear Ana, due 1 Nov."


def test_fill_missing_value(tmp_path):
    make_fill_template(tmp_path / "t_fill.docx")
    (tmp_path / "t_fill.json").write_text(json.dumps({"NAME": "Ana"}), encoding="utf-8")
    result = run_tool(tmp_path, "fill", "t_fill.docx", "--values", "t_fill.json")
    assert result.returncode == 1
    assert "DATE" in result.stderr
    assert not (tmp_path / "t_fill.fill.docx").exists()


def test_track_show(tmp_path):
    make_track_doc(tmp_path / "t_track.docx")
    result = run_tool(tmp_path, "track-show", "--json", "t_track.docx")
    assert result.returncode == 0, result.stderr
    rows = json.loads(result.stdout)["revisions"]
    assert [(r["kind"], r["author"], r["text"]) for r in rows] == [
        ("insert", "Ana", "new"),
        ("delete", "Ben", "old"),
    ]


def test_track_resolve_accept(tmp_path):
    make_track_doc(tmp_path / "t_track.docx")
    result = run_tool(tmp_path, "track-resolve", "t_track.docx", "--mode", "accept", "--all", "--json")
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["resolved"] == 2
    assert data["remaining"] == 0
    out = tmp_path / "t_track.track-resolve.docx"
    with __import__("zipfile").ZipFile(out) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    assert len(root.findall(f".//{{{W_NS}}}ins")) == 0
    assert len(root.findall(f".//{{{W_NS}}}del")) == 0
    assert read_json(tmp_path, out)["paragraphs"][0]["text"] == "new"


def test_track_resolve_unsupported(tmp_path):
    doc = Document()
    para = doc.add_paragraph("plain")
    change = OxmlElement("w:pPrChange")
    change.set(qn("w:id"), "7")
    change.set(qn("w:author"), "Ana")
    change.set(qn("w:date"), "2026-10-01T09:00:00Z")
    change.append(OxmlElement("w:pPr"))
    para._p.get_or_add_pPr().append(change)
    doc.save(tmp_path / "t_fmt.docx")
    result = run_tool(tmp_path, "track-resolve", "t_fmt.docx", "--mode", "accept", "--all")
    assert result.returncode == 2
    assert not (tmp_path / "t_fmt.track-resolve.docx").exists()


def _mark(tag, rid):
    """A paragraph-mark revision marker (w:ins or w:del with no content) for w:pPr/w:rPr."""
    el = OxmlElement(f"w:{tag}")
    el.set(qn("w:id"), str(rid))
    el.set(qn("w:author"), "Ana")
    el.set(qn("w:date"), "2026-10-01T09:00:00Z")
    return el


def _pm_para(doc, text="", *, style=None, mark=None, mark_id=None, content=None, content_id=None,
             section=False, change_id=None):
    """Append a paragraph. `mark` is "ins" or "del" for its paragraph mark; `content` is "ins" or
    "del" to wrap the text in a tracked revision; `section` adds a section break; `change_id`
    adds a paragraph-property change (w:pPrChange) whose old properties are empty."""
    p = doc.add_paragraph()
    if style:
        p.style = style
    ppr = p._p.get_or_add_pPr()
    if mark:
        rpr = OxmlElement("w:rPr")
        rpr.append(_mark(mark, mark_id))
        ppr.append(rpr)
    if section:
        ppr.append(OxmlElement("w:sectPr"))
    if change_id is not None:
        change = _mark("pPrChange", change_id)
        change.append(OxmlElement("w:pPr"))
        ppr.append(change)
    if text and content:
        p._p.append(_revision(content, content_id, "Ana", text, deleted=(content == "del")))
    elif text:
        run = OxmlElement("w:r")
        node = OxmlElement("w:t")
        node.text = text
        node.set(qn("xml:space"), "preserve")
        run.append(node)
        p._p.append(run)
    return p


def _resolve(tmp_path, src, mode, *extra):
    out = tmp_path / f"{mode}-{src.stem}.docx"
    result = run_tool(tmp_path, "track-resolve", src, "--mode", mode, *extra, "--out", out)
    return result, out


def _paras(tmp_path, path):
    return [(p["style"], p["text"]) for p in read_json(tmp_path, path)["paragraphs"]]


def _assert_no_tracked_markup(path):
    with __import__("zipfile").ZipFile(path) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    assert root.findall(f".//{{{W_NS}}}ins") == []
    assert root.findall(f".//{{{W_NS}}}del") == []


def test_pm_insert_middle_accept_and_reject(tmp_path):
    doc = Document()
    _pm_para(doc, "Alpha")
    _pm_para(doc, "New", mark="ins", mark_id=1, content="ins", content_id=2)
    _pm_para(doc, "Beta")
    src = tmp_path / "pm_mid.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Normal", "New"), ("Normal", "Beta")]
    _assert_no_tracked_markup(out)
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Normal", "Beta")]
    _assert_no_tracked_markup(out)


def test_pm_insert_start_and_end_typed(tmp_path):
    doc = Document()
    _pm_para(doc, "Start", mark="ins", mark_id=1, content="ins", content_id=2)
    _pm_para(doc, "Alpha")
    _pm_para(doc, "Beta", mark="ins", mark_id=3)
    _pm_para(doc, "End", content="ins", content_id=4)
    src = tmp_path / "pm_edges.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Start"), ("Normal", "Alpha"),
                                     ("Normal", "Beta"), ("Normal", "End")]
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Normal", "Beta")]
    _assert_no_tracked_markup(out)


def test_pm_insert_several_typed_reject(tmp_path):
    doc = Document()
    _pm_para(doc, "Alpha", mark="ins", mark_id=1)
    _pm_para(doc, "N1", mark="ins", mark_id=2, content="ins", content_id=3)
    _pm_para(doc, "N2", content="ins", content_id=4)
    _pm_para(doc, "Beta")
    src = tmp_path / "pm_several.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Normal", "N1"),
                                     ("Normal", "N2"), ("Normal", "Beta")]
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Normal", "Beta")]


def test_pm_whole_paragraph_delete_accept_and_reject(tmp_path):
    doc = Document()
    _pm_para(doc, "Alpha")
    _pm_para(doc, "Old heading", style="Heading 1", mark="del", mark_id=1, content="del", content_id=2)
    _pm_para(doc, "Old body", mark="del", mark_id=3, content="del", content_id=4)
    _pm_para(doc, "Gamma")
    src = tmp_path / "pm_whole_del.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Normal", "Gamma")]
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Heading 1", "Old heading"),
                                     ("Normal", "Old body"), ("Normal", "Gamma")]
    _assert_no_tracked_markup(out)


def test_pm_mark_delete_keeps_first_paragraph_style(tmp_path):
    """Word: accepting a deleted mark after a paragraph with text gives the merged paragraph the
    first paragraph's style."""
    doc = Document()
    _pm_para(doc, "Title", style="Heading 1", mark="del", mark_id=1)
    _pm_para(doc, "Body")
    src = tmp_path / "pm_merge.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Heading 1", "TitleBody")]
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Heading 1", "Title"), ("Normal", "Body")]


def test_pm_reject_inserted_mark_keeps_next_style(tmp_path):
    """Word: rejecting an inserted mark restores the next paragraph's style."""
    doc = Document()
    _pm_para(doc, "Section", style="Heading 1", mark="ins", mark_id=1)
    _pm_para(doc, " one")
    src = tmp_path / "pm_split.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Section one")]
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Heading 1", "Section"), ("Normal", " one")]


def test_pm_linked_format_change_accept_and_reject(tmp_path):
    """The paragraph after a deleted mark carries a formatting change (as Word writes it).
    Accept keeps the new style, reject restores the old one, with --all."""
    doc = Document()
    _pm_para(doc, "Title", style="Heading 1", mark="del", mark_id=1)
    _pm_para(doc, "Body", style="Heading 1", change_id=2)
    src = tmp_path / "pm_linked.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Heading 1", "TitleBody")]
    _assert_no_tracked_markup(out)
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Heading 1", "Title"), ("Normal", "Body")]
    _assert_no_tracked_markup(out)


def test_pm_linked_format_change_needs_selection(tmp_path):
    doc = Document()
    _pm_para(doc, "Title", style="Heading 1", mark="del", mark_id=1)
    _pm_para(doc, "Body", style="Heading 1", change_id=2)
    src = tmp_path / "pm_linked_one.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "accept", "--id", "1")
    assert result.returncode == 2
    assert "must be selected too" in result.stderr
    assert not out.exists()


def test_pm_section_break_refused_on_reject(tmp_path):
    doc = Document()
    _pm_para(doc, "Keep", mark="ins", mark_id=1, section=True)
    _pm_para(doc, "After")
    src = tmp_path / "pm_section.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 2
    assert "section" in result.stderr
    assert not out.exists()
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Keep"), ("Normal", "After")]


def test_pm_last_paragraph_mark_refused_on_reject(tmp_path):
    doc = Document()
    _pm_para(doc, "Alpha")
    _pm_para(doc, "Beta", mark="ins", mark_id=1)
    src = tmp_path / "pm_last.docx"
    doc.save(src)
    result, out = _resolve(tmp_path, src, "reject", "--all")
    assert result.returncode == 2
    assert "no paragraph after it" in result.stderr
    assert not out.exists()
    result, out = _resolve(tmp_path, src, "accept", "--all")
    assert result.returncode == 0, result.stderr
    assert _paras(tmp_path, out) == [("Normal", "Alpha"), ("Normal", "Beta")]


def test_pm_track_show_lists_paragraph_marks(tmp_path):
    doc = Document()
    _pm_para(doc, "Alpha")
    _pm_para(doc, "New", mark="ins", mark_id=1, content="ins", content_id=2)
    _pm_para(doc, "Beta")
    src = tmp_path / "pm_show.docx"
    doc.save(src)
    result = run_tool(tmp_path, "track-show", "--json", src)
    assert result.returncode == 0, result.stderr
    listed = json.loads(result.stdout)["revisions"]
    assert sorted((r["kind"], r["detail"] or "") for r in listed) == [("insert", ""), ("insert", "paragraph mark")]


def test_comments_add(tmp_path):
    make_replace_doc(tmp_path / "t_rep.docx")
    result = run_tool(tmp_path, "comments-add", "t_rep.docx", "--anchor", "signed",
                      "--text", "Check date", "--author", "Test Bot", "--json")
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["anchor_text"] == "signed"
    listed = run_tool(tmp_path, "comments-list", "--json", "t_rep.comments-add.docx")
    assert listed.returncode == 0, listed.stderr
    comments = json.loads(listed.stdout)["comments"]
    assert len(comments) == 1
    assert comments[0]["author"] == "Test Bot"
    assert comments[0]["text"] == "Check date"
    assert comments[0]["anchor_text"] == "signed"


def test_comments_add_ambiguous(tmp_path):
    doc = Document()
    doc.add_paragraph("ACME and ACME")
    doc.save(tmp_path / "t_amb.docx")
    result = run_tool(tmp_path, "comments-add", "t_amb.docx", "--anchor", "ACME",
                      "--text", "x", "--author", "Test Bot")
    assert result.returncode == 1
    assert "2" in result.stderr
    assert not (tmp_path / "t_amb.comments-add.docx").exists()


def test_validate_valid(tmp_path):
    make_summary_doc(tmp_path / "t_sum.docx")
    result = run_tool(tmp_path, "validate", "t_sum.docx")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "valid" in result.stdout.splitlines()[-1]


def test_validate_corrupt(tmp_path):
    source = make_summary_doc(tmp_path / "t_sum.docx")
    corrupt = tmp_path / "t_cut.docx"
    corrupt.write_bytes(source.read_bytes()[:1000])
    result = run_tool(tmp_path, "validate", "t_cut.docx")
    assert result.returncode == 1
    assert "file is not a readable zip" in result.stdout


def _libreoffice_present() -> bool:
    if shutil.which("soffice") or shutil.which("soffice.exe"):
        return True
    return Path(r"C:\Program Files\LibreOffice\program\soffice.exe").is_file() \
        or Path("/usr/bin/soffice").is_file() \
        or Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").is_file()


def _word_present() -> bool:
    """Word is registered and pywin32 can be imported by this Python (to-pdf needs both)."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            pass
    except (ImportError, OSError):
        return False
    return importlib.util.find_spec("win32com") is not None


def test_to_pdf(tmp_path):
    (tmp_path / "t_outline.txt").write_text("# Title\n\nHello world\n- one\n- two\n", encoding="utf-8")
    created = run_tool(tmp_path, "create", "t_outline.txt", "--out", "t_out.docx")
    assert created.returncode == 0, created.stderr
    result = run_tool(tmp_path, "to-pdf", "t_out.docx", "--out", "t_out.pdf")
    if _word_present() or _libreoffice_present():
        assert result.returncode == 0, result.stderr
        assert len(PdfReader(str(tmp_path / "t_out.pdf")).pages) == 1
    else:
        assert result.returncode == 2
        assert "LibreOffice is not installed. Install it or convert the file in Word." in result.stderr
        assert not (tmp_path / "t_out.pdf").exists()
