#!/usr/bin/env python3
"""Build the input fixtures and the task manifest (tasks.json) for the PDF skill benchmark.

Ground truth rule: every expected value below is written by hand from
specs/pdf.md and SKILL.md. Nothing in this file runs pdf_tool.py, and no
expected value is copied from a tool output.

Run:  <venv python> make_tasks.py
"""
import json
from io import BytesIO
from pathlib import Path

from pypdf import PasswordType, PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

HERE = Path(__file__).resolve().parent
FIX = HERE / "fixtures"
MANIFEST = HERE / "tasks.json"

PW = "t3st"  # password for enc.pdf (test value only)

# Phrases used across tasks. They are fixed text inside the fixtures.
Q_P1 = ["Northwind Traders - Quarterly Report, Q3 2026", "Revenue summary for July.", "Sales grew across all three regions."]
Q_P2 = ["Operating costs were reviewed in August.", "Freight spend fell by four percent."]
Q_P3 = ["Outlook for the fourth quarter is stable.", "Warehouse decision due in November."]


# ---------------------------------------------------------------- fixtures

def draw_lines(c, lines, top=750, step=24):
    c.setFont("Helvetica", 12)
    y = top
    for line in lines:
        c.drawString(72, y, line)
        y -= step


def make_text_pdf(path, pages):
    c = canvas.Canvas(str(path), pagesize=A4)
    for lines in pages:
        draw_lines(c, lines)
        c.showPage()
    c.save()


def make_invoice_pdf(path):
    styles = getSampleStyleSheet()
    data = [["Item", "Qty", "Price"], ["Widget", "2", "9.50"], ["Gasket", "10", "0.75"]]
    table = Table(data)
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    story = [
        Paragraph("Invoice INV-2026-0042", styles["Heading1"]),
        Paragraph("Bill to: Contoso Ltd", styles["Normal"]),
        Paragraph("Payment terms: 30 days", styles["Normal"]),
        PageBreak(),
        Paragraph("Line items", styles["Heading2"]),
        Spacer(1, 12),
        table,
    ]
    SimpleDocTemplate(str(path), pagesize=A4).build(story)


def make_scan_like_pdf(path):
    """Vector shapes only, no text layer: stands in for a scanned page."""
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFillGray(0.85)
    c.rect(60, 200, 480, 600, fill=1, stroke=0)
    c.setFillGray(0.6)
    c.rect(90, 300, 420, 40, fill=1, stroke=0)
    c.showPage()
    c.save()


def make_form_pdf(path):
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFont("Helvetica", 12)
    c.drawString(72, 740, "Customer sign-up form")
    c.drawString(72, 722, "Full name")
    c.acroForm.textfield(name="FullName", tooltip="Full name", x=72, y=690, width=250, height=20,
                         value="", borderStyle="inset", forceBorder=True)
    c.drawString(72, 662, "I agree to the terms")
    c.acroForm.checkbox(name="Agree", tooltip="Agree", x=72, y=630, size=14, buttonStyle="check",
                        checked=False, forceBorder=True)
    c.drawString(72, 602, "Plan")
    # reportlab 5.1 fails on a plain string value with options; a list value works and sets the default to "Basic".
    c.acroForm.choice(name="Plan", tooltip="Plan", value=["Basic"], x=72, y=570, width=150, height=20,
                      options=["Basic", "Pro", "Enterprise"], forceBorder=True, fieldFlags="combo")
    c.showPage()
    c.save()


def make_encrypted_copy(src, dst):
    reader = PdfReader(str(src))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(user_password=PW, owner_password="owner-not-used", algorithm="AES-256")
    with open(dst, "wb") as f:
        writer.write(f)


def make_stamp_pdf(path):
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setStrokeColor(colors.red)
    c.rect(360, 700, 160, 50, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 16)
    c.setFillColor(colors.red)
    c.drawString(380, 715, "APPROVED")
    c.showPage()
    c.save()


def make_rot270_pdf(path):
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    writer.add_blank_page(width=612, height=792)
    writer.pages[0].rotate(270)
    with open(path, "wb") as f:
        writer.write(f)


def build_fixtures():
    FIX.mkdir(parents=True, exist_ok=True)
    quarterly = FIX / "quarterly.pdf"
    make_text_pdf(quarterly, [Q_P1, Q_P2, Q_P3])
    make_invoice_pdf(FIX / "invoice.pdf")
    make_scan_like_pdf(FIX / "scan_like.pdf")
    make_form_pdf(FIX / "form.pdf")
    make_encrypted_copy(quarterly, FIX / "enc.pdf")
    make_text_pdf(FIX / "five_pages.pdf",
                  [[f"Section {i} of five", "Body text for this section."] for i in range(1, 6)])
    make_stamp_pdf(FIX / "stamp.pdf")
    make_rot270_pdf(FIX / "rot270.pdf")
    (FIX / "truncated.pdf").write_bytes(quarterly.read_bytes()[:500])
    (FIX / "empty.pdf").write_bytes(b"")
    (FIX / "existing.pdf").write_bytes(b"existing output, do not replace\n")


def validate_fixtures():
    """Fixture sanity only: confirms the inputs match the facts the tasks assume."""
    q = PdfReader(str(FIX / "quarterly.pdf"))
    assert len(q.pages) == 3
    assert "Revenue summary for July." in q.pages[0].extract_text()
    assert "Operating costs were reviewed in August." in q.pages[1].extract_text()
    assert "Outlook for the fourth quarter is stable." in q.pages[2].extract_text()
    inv = PdfReader(str(FIX / "invoice.pdf"))
    assert len(inv.pages) == 2 and "Invoice INV-2026-0042" in inv.pages[0].extract_text()
    scan = PdfReader(str(FIX / "scan_like.pdf"))
    assert not (scan.pages[0].extract_text() or "").strip()
    enc = PdfReader(str(FIX / "enc.pdf"))
    assert enc.is_encrypted and enc.decrypt(PW) != PasswordType.NOT_DECRYPTED and len(enc.pages) == 3
    form = PdfReader(str(FIX / "form.pdf"))
    assert set(form.get_fields().keys()) == {"FullName", "Agree", "Plan"}
    five = PdfReader(str(FIX / "five_pages.pdf"))
    assert len(five.pages) == 5
    rot = PdfReader(str(FIX / "rot270.pdf"))
    assert int(rot.pages[0].get("/Rotate", 0)) == 270 and len(rot.pages) == 2
    assert (FIX / "empty.pdf").stat().st_size == 0
    assert (FIX / "truncated.pdf").stat().st_size == 500


# ---------------------------------------------------------------- task text

NOTES_TXT = "Hello world\nSecond line\n\nThird line for the report.\n"
UNICODE_TXT = "Quarterly notes\nRevenue is up\n中文 notes\n"  # line 3 starts with U+4E2D
EXPECT_MIXED = "Revenue summary for July.\nBoard approval pending\noperating costs\n"
VALUES_OK = '{"FullName": "Ana", "Agree": true, "Plan": "Pro"}\n'
VALUES_UNKNOWN = '{"FullName": "Ana", "Nope": "x"}\n'
VALUES_EMPTY = "{}\n"


def tasks():
    T = []

    def add(**kw):
        kw.setdefault("needs", [])
        kw.setdefault("copy", {})
        kw.setdefault("write", {})
        kw.setdefault("env", {})
        kw.setdefault("checks", [])
        T.append(kw)

    q = {"quarterly.pdf": "quarterly.pdf"}

    # ---- text
    add(id="01_text_all_pages_json", category="text", kind="happy",
        request="Pull the text out of the Q3 report so I can paste it into the summary.",
        copy=q, args=["text", "quarterly.pdf", "--json"], exit=0,
        checks=[
            {"t": "json_equals", "map": {"page_count": 3, "pages.len": 3, "no_text_pages": [],
                                          "pages.0.page": 1, "pages.1.page": 2, "pages.2.page": 3}},
            {"t": "json_contains", "map": {"pages.0.text": ["Revenue summary for July."],
                                           "pages.1.text": ["Operating costs were reviewed in August."],
                                           "pages.2.text": ["Outlook for the fourth quarter is stable."]}},
        ])

    add(id="02_text_scanned_refused", category="text", kind="refusal",
        request="Get the text from this scanned contract scan.",
        copy={"scan_like.pdf": "scan_like.pdf"}, args=["text", "scan_like.pdf"], exit=2,
        checks=[{"t": "stderr_contains", "items": ["No text layer", "OCR"]}])

    add(id="03_text_page_list_out_of_range", category="text", kind="happy",
        request="Give me pages 2 to 9 of the Q3 report as text.",
        copy=q, args=["text", "quarterly.pdf", "--pages", "2-9", "--json"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["Page list is invalid or outside 1..3"]},
                {"t": "stdout_empty"}])

    # ---- password
    add(id="04_password_no_env_refused", category="password", kind="safety",
        request="Read the locked Q3 report. I have not set up the password.",
        copy={"enc.pdf": "enc.pdf"}, args=["text", "enc.pdf"], exit=2,
        checks=[{"t": "stderr_contains", "items": ["password-protected", "PDF_PASSWORD"]},
                {"t": "no_secret", "values": [PW]}])

    add(id="05_password_correct_env", category="password", kind="happy",
        request="Open the locked Q3 report with the password in PDF_TEST_PW and extract the text.",
        copy={"enc.pdf": "enc.pdf"}, env={"PDF_TEST_PW": PW},
        args=["text", "enc.pdf", "--password-env", "PDF_TEST_PW", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"page_count": 3}},
                {"t": "json_contains", "map": {"pages.0.text": ["Revenue summary for July."]}},
                {"t": "no_secret", "values": [PW]}])

    add(id="06_password_wrong_no_leak", category="password", kind="safety",
        request="Open the locked Q3 report with the password in PDF_TEST_PW (it may be wrong).",
        copy={"enc.pdf": "enc.pdf"}, env={"PDF_TEST_PW": "not-the-pw"},
        args=["text", "enc.pdf", "--password-env", "PDF_TEST_PW", "--json"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["did not open this PDF"]},
                {"t": "no_secret", "values": ["not-the-pw"]}])

    # ---- tables
    add(id="07_tables_invoice_rows", category="tables", kind="happy",
        request="Pull the line items table out of invoice INV-2026-0042.",
        copy={"invoice.pdf": "invoice.pdf"}, args=["tables", "invoice.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {
            "tables.len": 1, "tables.0.page": 2, "tables.0.index": 1,
            "tables.0.rows": [["Item", "Qty", "Price"], ["Widget", "2", "9.50"], ["Gasket", "10", "0.75"]]}}])

    # ---- fields
    add(id="08_fields_list_form", category="fields", kind="happy",
        request="What fields does this sign-up form have, and what are the Plan options?",
        copy={"form.pdf": "form.pdf"}, args=["fields-list", "form.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"fields.len": 3}},
                {"t": "json_items", "path": "fields", "key": "name", "expect": [
                    {"name": "FullName", "type": "text", "value": "", "options": [], "page": 1},
                    {"name": "Agree", "type": "checkbox", "value": False, "options": [], "page": 1},
                    {"name": "Plan", "type": "dropdown", "value": "Basic", "options": ["Basic", "Pro", "Enterprise"],
                     "page": 1}]}])

    add(id="09_fields_fill_ok", category="fields", kind="happy",
        request="Fill in Ana's sign-up: name Ana, agrees to terms, plan Pro.",
        copy={"form.pdf": "form.pdf"}, write={"values_ok.json": VALUES_OK},
        args=["fields-fill", "form.pdf", "--values", "values_ok.json", "--out", "filled.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"verified": True,
                                              "filled": {"FullName": "Ana", "Agree": True, "Plan": "Pro"}}},
                {"t": "pdf", "path": "filled.pdf", "page_count": 1,
                 "fields": {"FullName": "Ana", "Agree": True, "Plan": "Pro"}},
                {"t": "file_unchanged", "path": "form.pdf"}])

    add(id="10_fields_fill_unknown_no_write", category="fields", kind="refusal",
        request="Fill the sign-up form with name Ana and a field called Nope.",
        copy={"form.pdf": "form.pdf"}, write={"values_unknown.json": VALUES_UNKNOWN},
        args=["fields-fill", "form.pdf", "--values", "values_unknown.json", "--out", "filled_bad.pdf"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["Nope", "No file written"]},
                {"t": "file_exists", "path": "filled_bad.pdf", "exists": False}])

    add(id="11_fields_fill_no_form", category="fields", kind="refusal",
        request="Fill the form fields in the Q3 report.",
        copy=q, write={"values_empty.json": VALUES_EMPTY},
        args=["fields-fill", "quarterly.pdf", "--values", "values_empty.json", "--out", "filled_none.pdf"], exit=2,
        checks=[{"t": "stderr_contains", "items": ["This PDF has no form fields."]},
                {"t": "file_exists", "path": "filled_none.pdf", "exists": False}])

    # ---- merge
    add(id="12_merge_two_pdfs", category="merge", kind="happy",
        request="Combine the Q3 report and invoice INV-2026-0042 into one PDF, report first.",
        copy={"quarterly.pdf": "quarterly.pdf", "invoice.pdf": "invoice.pdf"},
        args=["merge", "quarterly.pdf", "invoice.pdf", "--out", "merged.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"verified": True}},
                {"t": "pdf", "path": "merged.pdf", "page_count": 5,
                 "page_text_contains": {"1": ["Revenue summary for July."],
                                        "4": ["Invoice INV-2026-0042"],
                                        "5": ["Widget"]}}])

    add(id="13_merge_output_exists_no_force", category="merge", kind="safety",
        request="Merge the report and invoice into existing.pdf.",
        copy={"quarterly.pdf": "quarterly.pdf", "invoice.pdf": "invoice.pdf", "existing.pdf": "existing.pdf"},
        args=["merge", "quarterly.pdf", "invoice.pdf", "--out", "existing.pdf"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["Output already exists"]},
                {"t": "file_unchanged", "path": "existing.pdf"}])

    # ---- split
    add(id="14_split_every_2", category="split", kind="happy",
        request="Split this 5-page document into parts of 2 pages each.",
        copy={"five_pages.pdf": "five_pages.pdf"},
        args=["split", "five_pages.pdf", "--every", "2", "--out-dir", "parts"], exit=0,
        checks=[{"t": "pdf", "path": "parts/five_pages.part01.pdf", "page_count": 2,
                 "page_text_sequence": {"1": ["Section 1 of five"], "2": ["Section 2 of five"]}},
                {"t": "pdf", "path": "parts/five_pages.part02.pdf", "page_count": 2,
                 "page_text_sequence": {"1": ["Section 3 of five"], "2": ["Section 4 of five"]}},
                {"t": "pdf", "path": "parts/five_pages.part03.pdf", "page_count": 1,
                 "page_text_sequence": {"1": ["Section 5 of five"]}},
                {"t": "file_exists", "path": "parts/five_pages.part04.pdf", "exists": False}])

    add(id="15_split_ranges_gap", category="split", kind="refusal",
        request="Split the 5-page document into pages 1-2 and 4-5, leaving out page 3.",
        copy={"five_pages.pdf": "five_pages.pdf"},
        args=["split", "five_pages.pdf", "--ranges", "1-2,4-5", "--out-dir", "parts_gap"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["Page 3 is not covered"]},
                {"t": "no_pdf_in", "path": "parts_gap"}])

    # ---- rotate
    add(id="16_rotate_page_90", category="rotate", kind="happy",
        request="Rotate page 1 of the Q3 report 90 degrees clockwise.",
        copy=q, args=["rotate", "quarterly.pdf", "--pages", "1", "--angle", "90", "--out", "rotated.pdf", "--json"],
        exit=0,
        checks=[{"t": "json_equals", "map": {"verified": True}},
                {"t": "pdf", "path": "rotated.pdf", "page_count": 3, "rotations": {"1": 90, "2": 0, "3": 0},
                 "page_text_contains": {"1": ["Revenue summary for July."]}},
                {"t": "file_unchanged", "path": "quarterly.pdf"}])

    add(id="17_rotate_wraps_360", category="rotate", kind="happy",
        request="Page 1 of this scan is already turned 270 degrees. Turn it another 180.",
        copy={"rot270.pdf": "rot270.pdf"},
        args=["rotate", "rot270.pdf", "--pages", "1", "--angle", "180", "--out", "rot_wrap.pdf"], exit=0,
        checks=[{"t": "pdf", "path": "rot_wrap.pdf", "page_count": 2, "rotations": {"1": 90, "2": 0}}])

    add(id="18_rotate_input_overwrite_refused", category="rotate", kind="safety",
        request="Rotate the Q3 report and save it over the original file.",
        copy=q, args=["rotate", "quarterly.pdf", "--angle", "90", "--out", "quarterly.pdf"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["Refusing to overwrite an input"]},
                {"t": "file_unchanged", "path": "quarterly.pdf"}])

    add(id="19_rotate_missing_angle", category="rotate", kind="refusal",
        request="Rotate the Q3 report and save it as no_angle.pdf.",
        copy=q, args=["rotate", "quarterly.pdf", "--out", "no_angle.pdf"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["--angle"]},
                {"t": "file_exists", "path": "no_angle.pdf", "exists": False}])

    # ---- reorder
    add(id="20_reorder_3_1_2", category="reorder", kind="happy",
        request="Move the outlook page of the Q3 report to the front, then revenue, then costs.",
        copy=q, args=["reorder", "quarterly.pdf", "--order", "3,1,2", "--out", "reordered.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"verified": True}},
                {"t": "pdf", "path": "reordered.pdf", "page_count": 3,
                 "page_text_sequence": {"1": ["Outlook for the fourth quarter is stable."],
                                        "2": ["Revenue summary for July."],
                                        "3": ["Operating costs were reviewed in August."]}}])

    add(id="21_reorder_not_permutation", category="reorder", kind="refusal",
        request="Reorder the Q3 report to pages 1 and 2 only.",
        copy=q, args=["reorder", "quarterly.pdf", "--order", "1,2", "--out", "bad_order.pdf"], exit=1,
        checks=[{"t": "stderr_contains", "items": ["Order must list every page 1..3 exactly once."]},
                {"t": "file_exists", "path": "bad_order.pdf", "exists": False}])

    # ---- watermark
    add(id="22_watermark_stamp", category="watermark", kind="happy",
        request="Stamp APPROVED on every page of the Q3 report.",
        copy={"quarterly.pdf": "quarterly.pdf", "stamp.pdf": "stamp.pdf"},
        args=["watermark", "quarterly.pdf", "--stamp", "stamp.pdf", "--out", "stamped.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"verified": True}},
                {"t": "pdf", "path": "stamped.pdf", "page_count": 3,
                 "page_text_contains": {"1": ["Revenue summary for July."],
                                        "3": ["Outlook for the fourth quarter is stable."]}}])

    add(id="23_watermark_text_draft", category="watermark", kind="happy",
        request="Put a DRAFT watermark on the Q3 report.",
        copy=q, args=["watermark", "quarterly.pdf", "--text", "DRAFT", "--out", "draft.pdf", "--json"], exit=0,
        checks=[{"t": "pdf", "path": "draft.pdf", "page_count": 3,
                 "page_text_contains": {"1": ["Revenue summary for July.", "DRAFT"],
                                        "2": ["Operating costs were reviewed in August.", "DRAFT"]}}])

    # ---- create
    add(id="24_create_basic", category="create", kind="happy",
        request="Turn my meeting notes into a PDF.",
        write={"notes.txt": NOTES_TXT},
        args=["create", "notes.txt", "--out", "notes.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"verified": True}},
                {"t": "pdf", "path": "notes.pdf", "page_count": 1,
                 "page_text_sequence": {"1": ["Hello world", "Second line", "Third line for the report."]}}])

    add(id="25_create_unsupported_char", category="create", kind="refusal",
        request="Make a PDF from these notes (they contain a Chinese word).",
        write={"unicode_notes.txt": UNICODE_TXT},
        args=["create", "unicode_notes.txt", "--out", "cjk.pdf"], exit=2,
        checks=[{"t": "stderr_contains", "items": ["U+4E2D", "line 3"]},
                {"t": "file_exists", "path": "cjk.pdf", "exists": False}])

    # ---- check
    add(id="26_check_valid", category="check", kind="happy",
        request="Is this Q3 report a clean, readable PDF?",
        copy=q, args=["check", "quarterly.pdf", "--json"], exit=0,
        checks=[{"t": "json_equals", "map": {"readable": True, "page_count": 3, "encrypted": False, "errors": []}}])

    add(id="27_check_truncated_corrupt", category="check", kind="safety",
        request="The upload of the Q3 report looks cut off. Is it readable?",
        copy={"truncated.pdf": "truncated.pdf"}, args=["check", "truncated.pdf", "--json"], exit=1,
        checks=[{"t": "json_equals", "map": {"readable": False}},
                {"t": "json_len_min", "path": "errors", "min": 1}])

    add(id="28_check_empty_file", category="check", kind="safety",
        request="Check this 0-byte download named report.pdf.",
        copy={"empty.pdf": "empty.pdf"}, args=["check", "empty.pdf", "--json"], exit=1,
        checks=[{"t": "json_equals", "map": {"readable": False}},
                {"t": "json_len_min", "path": "errors", "min": 1}])

    add(id="29_check_encrypted_no_pw", category="check", kind="refusal",
        request="Check whether the locked Q3 report is readable.",
        copy={"enc.pdf": "enc.pdf"}, args=["check", "enc.pdf", "--json"], exit=2,
        checks=[{"t": "stderr_contains", "items": ["password-protected"]},
                {"t": "no_secret", "values": [PW]}])

    # ---- check-text
    add(id="30_checktext_missing_and_case", category="check-text", kind="happy",
        request="Confirm the Q3 report contains these three phrases before I export it.",
        copy=q, write={"expect_mixed.txt": EXPECT_MIXED},
        args=["check-text", "quarterly.pdf", "--expect", "expect_mixed.txt", "--json"], exit=1,
        checks=[{"t": "json_equals", "map": {
            "found": [{"phrase": "Revenue summary for July.", "page": 1}],
            "missing": ["Board approval pending", "operating costs"]}},
                {"t": "stderr_contains", "items": ["2 phrase(s) missing"]}])

    return T


def main():
    build_fixtures()
    validate_fixtures()
    T = tasks()
    ids = [t["id"] for t in T]
    assert len(ids) == len(set(ids)), "duplicate task id"
    assert 25 <= len(T) <= 30, f"task count {len(T)} outside 25..30"
    for t in T:
        for name in t["copy"].values():
            assert (FIX / name).exists(), f"{t['id']}: missing fixture {name}"
    manifest = {
        "note": "Expected values are hand-written from specs/pdf.md and SKILL.md. Not taken from tool output.",
        "count": len(T),
        "tasks": T,
    }
    MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"wrote {len(T)} tasks to {MANIFEST}")


if __name__ == "__main__":
    main()
