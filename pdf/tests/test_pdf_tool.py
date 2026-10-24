"""Acceptance tests for pdf_tool.py, one test per acceptance test in specs/pdf.md section 3.

Every input file is built inside the test (reportlab or pypdf). The tool is run as a
separate process, the way an agent would run it.
"""
import io
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

if sys.platform == "win32":
    import msvcrt

TOOL = Path(__file__).resolve().parents[1] / "scripts" / "pdf_tool.py"
PASSWORD_VAR = "PDF_TEST_PW"
LOCKED_MESSAGE = "The file is open in another program. Close it and try again."


def run(*args, env_remove=(), env_set=None):
    env = dict(os.environ)
    for name in env_remove:
        env.pop(name, None)
    if env_set:
        env.update(env_set)
    return subprocess.run(
        [sys.executable, str(TOOL), *[str(a) for a in args]],
        capture_output=True, text=True, encoding="utf-8", env=env,
    )


def make_text_pdf(path: Path, pages):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(612, 792))
    for text in pages:
        c.drawString(72, 700, text)
        c.showPage()
    c.save()
    path.write_bytes(buf.getvalue())
    return path


def make_blank_pdf(path: Path, pages: int):
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(612, 792)
    with open(path, "wb") as fh:
        writer.write(fh)
    return path


def make_table_pdf(path: Path):
    doc = SimpleDocTemplate(str(path), pagesize=(612, 792))
    table = Table([["a", "b"], ["c", "d"]])
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    doc.build([table])
    return path


def make_form_pdf(path: Path):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(612, 792))
    c.drawString(72, 740, "Form")
    c.acroForm.textfield(name="FullName", tooltip="Full name", x=72, y=600,
                         width=200, height=20, value="", borderStyle="inset", forceBorder=True)
    c.acroForm.checkbox(name="Agree", tooltip="Agree", x=72, y=550, size=14, checked=False)
    c.showPage()
    c.save()
    path.write_bytes(buf.getvalue())
    return path


def make_encrypted_pdf(path: Path, user_password: str):
    plain = make_text_pdf(path.with_suffix(".plain.pdf"), ["Alpha", "Beta"])
    writer = PdfWriter(clone_from=PdfReader(str(plain)))
    writer.encrypt(user_password=user_password, owner_password="owner-" + user_password,
                   algorithm="AES-256")
    with open(path, "wb") as fh:
        writer.write(fh)
    plain.unlink()
    return path


def page_texts(path: Path):
    return [p.extract_text() or "" for p in PdfReader(str(path)).pages]


def test_text_extracts_three_pages_in_order(tmp_path):
    src = make_text_pdf(tmp_path / "t_text.pdf", ["Alpha", "Beta", "Gamma"])
    result = run("text", src, "--json")
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["page_count"] == 3
    texts = [p["text"] for p in data["pages"]]
    assert "Alpha" in texts[0]
    assert "Beta" in texts[1]
    assert "Gamma" in texts[2]


def test_tables_returns_the_two_by_two_table(tmp_path):
    src = make_table_pdf(tmp_path / "t_table.pdf")
    result = run("tables", src, "--json")
    assert result.returncode == 0, result.stderr
    tables = json.loads(result.stdout)["tables"]
    assert len(tables) == 1
    assert tables[0]["page"] == 1
    assert tables[0]["rows"] == [["a", "b"], ["c", "d"]]


def test_fields_list_reports_text_and_checkbox(tmp_path):
    src = make_form_pdf(tmp_path / "t_form.pdf")
    result = run("fields-list", src, "--json")
    assert result.returncode == 0, result.stderr
    rows = {row["name"]: row for row in json.loads(result.stdout)["fields"]}
    assert len(rows) == 2
    assert rows["FullName"]["type"] == "text"
    assert rows["Agree"]["type"] == "checkbox"


def test_fields_fill_writes_verified_values_and_refuses_unknown_names(tmp_path):
    src = make_form_pdf(tmp_path / "t_form.pdf")
    values = tmp_path / "values.json"
    values.write_text(json.dumps({"FullName": "Ana", "Agree": True}), encoding="utf-8")
    out = tmp_path / "filled.pdf"
    result = run("fields-fill", src, "--values", values, "--out", out)
    assert result.returncode == 0, result.stderr
    listed = run("fields-list", out, "--json")
    rows = {row["name"]: row for row in json.loads(listed.stdout)["fields"]}
    assert rows["FullName"]["value"] == "Ana"
    assert rows["Agree"]["value"] is True

    bad_values = tmp_path / "bad.json"
    bad_values.write_text(json.dumps({"Nope": "x"}), encoding="utf-8")
    bad_out = tmp_path / "nope.pdf"
    result = run("fields-fill", src, "--values", bad_values, "--out", bad_out)
    assert result.returncode == 1
    assert "Nope" in result.stderr
    assert not bad_out.exists()


def test_merge_joins_blank_pdfs_in_order(tmp_path):
    a = make_blank_pdf(tmp_path / "a.pdf", 2)
    b = make_blank_pdf(tmp_path / "b.pdf", 3)
    out = tmp_path / "m.pdf"
    result = run("merge", a, b, "--out", out)
    assert result.returncode == 0, result.stderr
    assert len(PdfReader(str(out)).pages) == 5


def test_split_every_two_makes_three_parts(tmp_path):
    src = make_blank_pdf(tmp_path / "five.pdf", 5)
    folder = tmp_path / "parts"
    result = run("split", src, "--every", 2, "--out-dir", folder)
    assert result.returncode == 0, result.stderr
    counts = [len(PdfReader(str(folder / name)).pages) for name in
              ("five.part01.pdf", "five.part02.pdf", "five.part03.pdf")]
    assert counts == [2, 2, 1]


def test_split_ranges_with_gap_is_refused(tmp_path):
    src = make_blank_pdf(tmp_path / "five.pdf", 5)
    folder = tmp_path / "gap_parts"
    result = run("split", src, "--ranges", "1-2,4-5", "--out-dir", folder)
    assert result.returncode == 1
    assert "Page 3 is not covered" in result.stderr
    assert not folder.exists()


def test_rotate_sets_rotation_on_page_one_only(tmp_path):
    src = make_blank_pdf(tmp_path / "two.pdf", 2)
    out = tmp_path / "rotated.pdf"
    result = run("rotate", src, "--pages", 1, "--angle", 90, "--out", out)
    assert result.returncode == 0, result.stderr
    pages = PdfReader(str(out)).pages
    assert pages[0].get("/Rotate") == 90
    assert pages[1].get("/Rotate", 0) == 0


def test_reorder_applies_permutation_and_rejects_partial_lists(tmp_path):
    src = make_text_pdf(tmp_path / "abg.pdf", ["Alpha", "Beta", "Gamma"])
    out = tmp_path / "reordered.pdf"
    result = run("reorder", src, "--order", "3,1,2", "--out", out)
    assert result.returncode == 0, result.stderr
    texts = page_texts(out)
    assert "Gamma" in texts[0]
    assert "Alpha" in texts[1]
    assert "Beta" in texts[2]

    partial = run("reorder", src, "--order", "1,2", "--out", tmp_path / "partial.pdf")
    assert partial.returncode == 1
    assert "Order must list every page 1..3 exactly once." in partial.stderr


def test_watermark_with_stamp_keeps_page_count(tmp_path):
    src = make_blank_pdf(tmp_path / "two.pdf", 2)
    stamp = make_blank_pdf(tmp_path / "stamp.pdf", 1)
    out = tmp_path / "watermarked.pdf"
    result = run("watermark", src, "--stamp", stamp, "--out", out)
    assert result.returncode == 0, result.stderr
    assert len(PdfReader(str(out)).pages) == 2


def test_create_makes_pdf_and_refuses_unsupported_character(tmp_path):
    text_file = tmp_path / "t_create.txt"
    text_file.write_text("Hello world\nSecond line\n", encoding="utf-8")
    out = tmp_path / "t_create.pdf"
    result = run("create", text_file, "--out", out)
    assert result.returncode == 0, result.stderr
    reader = PdfReader(str(out))
    assert len(reader.pages) == 1
    text = " ".join((reader.pages[0].extract_text() or "").split())
    assert text.index("Hello world") < text.index("Second line")

    bad_file = tmp_path / "bad.txt"
    bad_file.write_text("Chinese: 中\n", encoding="utf-8")
    bad_out = tmp_path / "bad.pdf"
    result = run("create", bad_file, "--out", bad_out)
    assert result.returncode == 2
    assert "U+4E2D" in result.stderr
    assert not bad_out.exists()


def test_check_passes_a_valid_pdf(tmp_path):
    src = make_text_pdf(tmp_path / "t_text.pdf", ["Alpha", "Beta", "Gamma"])
    result = run("check", src, "--json")
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["readable"] is True
    assert data["page_count"] == 3


def test_check_fails_a_truncated_pdf(tmp_path):
    full = make_text_pdf(tmp_path / "t_text.pdf", ["Alpha", "Beta", "Gamma"])
    src = tmp_path / "truncated.pdf"
    src.write_bytes(full.read_bytes()[:500])
    result = run("check", src, "--json")
    assert result.returncode == 1
    data = json.loads(result.stdout)
    assert data["readable"] is False
    assert len(data["errors"]) >= 1


def test_password_rules_for_encrypted_pdf(tmp_path):
    src = make_encrypted_pdf(tmp_path / "t_enc.pdf", "t3st")

    no_env = run("text", src, env_remove=(PASSWORD_VAR,))
    assert no_env.returncode == 2
    assert "password-protected" in no_env.stderr

    right = run("text", src, "--password-env", PASSWORD_VAR, "--json",
                env_set={PASSWORD_VAR: "t3st"})
    assert right.returncode == 0, right.stderr
    assert "Alpha" in json.loads(right.stdout)["pages"][0]["text"]
    assert "t3st" not in right.stdout

    wrong = run("text", src, "--password-env", PASSWORD_VAR, env_set={PASSWORD_VAR: "nope"})
    assert wrong.returncode == 1


def test_check_text_reports_found_and_missing_phrases(tmp_path):
    src = make_text_pdf(tmp_path / "alpha_beta.pdf", ["Alpha", "Beta"])
    expect = tmp_path / "expect.txt"
    expect.write_text("Alpha\nGamma\n", encoding="utf-8")
    result = run("check-text", src, "--expect", expect, "--json")
    assert result.returncode == 1
    data = json.loads(result.stdout)
    assert data["found"] == [{"phrase": "Alpha", "page": 1}]
    assert data["missing"] == ["Gamma"]
    assert "1 phrase(s) missing: Gamma" in result.stderr


def _folder_names(folder: Path):
    return sorted(p.name for p in folder.iterdir())


def test_in_place_rotate_replaces_input_and_leaves_no_temp_file(tmp_path):
    src = make_blank_pdf(tmp_path / "two.pdf", 2)
    result = run("rotate", src, "--pages", 1, "--angle", 90, "--in-place")
    assert result.returncode == 0, result.stderr
    pages = PdfReader(str(src)).pages
    assert pages[0].get("/Rotate") == 90
    assert pages[1].get("/Rotate", 0) == 0
    assert _folder_names(tmp_path) == ["two.pdf"]


def test_in_place_reorder_replaces_input(tmp_path):
    src = make_text_pdf(tmp_path / "abg.pdf", ["Alpha", "Beta", "Gamma"])
    result = run("reorder", src, "--order", "3,1,2", "--in-place")
    assert result.returncode == 0, result.stderr
    texts = page_texts(src)
    assert "Gamma" in texts[0]
    assert "Alpha" in texts[1]
    assert "Beta" in texts[2]
    assert _folder_names(tmp_path) == ["abg.pdf"]


def test_in_place_fields_fill_replaces_input(tmp_path):
    src = make_form_pdf(tmp_path / "t_form.pdf")
    values = tmp_path / "values.json"
    values.write_text(json.dumps({"FullName": "Ana", "Agree": True}), encoding="utf-8")
    result = run("fields-fill", src, "--values", values, "--in-place")
    assert result.returncode == 0, result.stderr
    rows = {row["name"]: row for row in json.loads(run("fields-list", src, "--json").stdout)["fields"]}
    assert rows["FullName"]["value"] == "Ana"
    assert rows["Agree"]["value"] is True


def test_in_place_watermark_keeps_original_text(tmp_path):
    src = make_text_pdf(tmp_path / "memo.pdf", ["Alpha"])
    result = run("watermark", src, "--text", "DRAFT", "--in-place")
    assert result.returncode == 0, result.stderr
    assert "Alpha" in page_texts(src)[0]
    assert _folder_names(tmp_path) == ["memo.pdf"]


def test_in_place_is_refused_for_encrypted_input_and_keeps_it_encrypted(tmp_path):
    src = make_encrypted_pdf(tmp_path / "t_enc.pdf", "t3st")
    before = src.read_bytes()
    result = run("rotate", src, "--angle", 90, "--in-place", "--password-env", PASSWORD_VAR,
                 env_set={PASSWORD_VAR: "t3st"})
    assert result.returncode == 1
    assert "--in-place is refused for a password-protected PDF" in result.stderr
    assert src.read_bytes() == before
    assert _folder_names(tmp_path) == ["t_enc.pdf"]


def test_in_place_and_out_are_exclusive_and_one_is_required(tmp_path):
    src = make_blank_pdf(tmp_path / "two.pdf", 2)
    both = run("rotate", src, "--angle", 90, "--out", tmp_path / "x.pdf", "--in-place")
    assert both.returncode == 1
    assert "not allowed with argument" in both.stderr
    neither = run("rotate", src, "--angle", 90)
    assert neither.returncode == 1
    assert "--in-place" in neither.stderr and "--out" in neither.stderr
    assert _folder_names(tmp_path) == ["two.pdf"]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows byte-range locks")
def test_locked_input_reports_open_in_another_program(tmp_path):
    src = make_text_pdf(tmp_path / "locked.pdf", ["Alpha"])
    before = src.read_bytes()
    size = os.path.getsize(src)
    handle = open(src, "r+b")
    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, size)
    try:
        in_place = run("rotate", src, "--angle", 90, "--in-place")
        reading = run("text", src)
    finally:
        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, size)
        handle.close()
    assert in_place.returncode == 1
    assert in_place.stderr.strip() == LOCKED_MESSAGE
    assert reading.returncode == 1
    assert reading.stderr.strip() == LOCKED_MESSAGE
    assert src.read_bytes() == before
    assert _folder_names(tmp_path) == ["locked.pdf"]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows file-sharing rules")
def test_open_handle_blocks_in_place_replace_with_locked_message(tmp_path):
    src = make_text_pdf(tmp_path / "held.pdf", ["Alpha"])
    before = src.read_bytes()
    with open(src, "rb") as handle:
        result = run("rotate", src, "--angle", 90, "--in-place")
        assert handle.read() == before
    assert result.returncode == 1
    assert result.stderr.strip() == LOCKED_MESSAGE
    assert src.read_bytes() == before
    assert _folder_names(tmp_path) == ["held.pdf"]
