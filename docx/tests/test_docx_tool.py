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
