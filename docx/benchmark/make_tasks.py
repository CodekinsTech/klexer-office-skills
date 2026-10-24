"""DOCX v2 accuracy benchmark: task definitions and input fixtures.

Expected results are written by hand from the v2 spec (specs/docx.md) and the
CLI rules in docx/SKILL.md. Nothing in this file runs docx_tool.py.

Input documents are built with python-docx. Tracked changes, property
revisions and comment anchors that python-docx has no API for are added as raw
OOXML elements. The fact readers at the bottom use python-docx, zipfile and
lxml only, so the checks do not depend on the tool under test.
"""

import io
import re
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

DATE = "2026-10-01T10:00:00Z"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
P_TAG = f"{{{W}}}p"
R_TAG = f"{{{W}}}r"
T_TAG = f"{{{W}}}t"
DT_TAG = f"{{{W}}}delText"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

FREE = 0xFFFFFFFF
END = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


def _cfb(streams):
    """Minimal OLE compound file (version 3, 512-byte sectors) holding the given streams.

    Each stream must be at least 4096 bytes and a multiple of 512 bytes, so it lives
    in regular sectors and no mini stream is needed. Written from the MS-CFB layout.
    """
    import struct
    streams = sorted(streams, key=lambda s: (len(s[0]), s[0].upper()))
    fat = [FATSECT, END]  # sector 0 holds the FAT, sector 1 the directory
    body = b""
    starts = []
    next_sector = 2
    for _name, data in streams:
        assert len(data) >= 4096 and len(data) % 512 == 0
        count = len(data) // 512
        starts.append((next_sector, len(data)))
        for i in range(count):
            fat.append(next_sector + 1 if i < count - 1 else END)
            next_sector += 1
        body += data
    assert len(fat) <= 128, "one FAT sector only"
    fat += [FREE] * (128 - len(fat))

    def entry(name, etype, left, right, child, start, size):
        raw = name.encode("utf-16-le") + b"\x00\x00"
        e = raw.ljust(64, b"\x00")
        e += struct.pack("<HBBIII", len(name) * 2 + 2, etype, 1, left, right, child)
        e += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        e += struct.pack("<IQ", start, size)
        return e

    dir_bytes = entry("Root Entry", 5, NOSTREAM, NOSTREAM, 1, END, 0)
    for i, ((name, _data), (start, size)) in enumerate(zip(streams, starts)):
        right = i + 2 if i + 1 < len(streams) else NOSTREAM
        dir_bytes += entry(name, 2, NOSTREAM, right, NOSTREAM, start, size)
    dir_bytes = dir_bytes.ljust(512, b"\x00")

    hdr = bytearray(512)
    hdr[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", hdr, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIII", hdr, 40, 0, 1, 1, 0, 4096)
    struct.pack_into("<III", hdr, 60, END, 0, END)
    struct.pack_into("<I", hdr, 72, 0)
    struct.pack_into("<I", hdr, 76, 0)
    for i in range(1, 109):
        struct.pack_into("<I", hdr, 76 + 4 * i, FREE)
    return bytes(hdr) + struct.pack("<128I", *fat) + dir_bytes + body


def _encryption_info():
    """Start of an EncryptionInfo stream: version 4.4 header, then filler."""
    import struct
    return (struct.pack("<HHI", 4, 4, 0x40) + bytes(range(256)) * 16)[:4096]


# Legacy .doc: an OLE file with a WordDocument stream and a .doc name.
# Encrypted .docx: an OLE file with EncryptionInfo and EncryptedPackage streams.
# Neither is a Word-generated file; both are structural stand-ins (see README).
def legacy_doc_bytes():
    return _cfb([("WordDocument", b"\x00" * 4096)])


def encrypted_docx_bytes():
    return _cfb([("EncryptionInfo", _encryption_info()),
                 ("EncryptedPackage", bytes(range(256)) * 16)])


# ---------------------------------------------------------------- builders

def _docx_bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _write_text(path, text, encoding="utf-8"):
    path.write_bytes(text.encode(encoding))


def _table(doc, rows):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            t.cell(r, c).text = value
    return t


def _run_el(text, deleted=False):
    r = OxmlElement("w:r")
    t = OxmlElement("w:delText" if deleted else "w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def _revision(par, tag, text, author, rev_id, deleted=False):
    el = OxmlElement(tag)
    el.set(qn("w:id"), str(rev_id))
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), DATE)
    el.append(_run_el(text, deleted=deleted))
    par._p.append(el)


def add_ins(par, text, author, rev_id):
    _revision(par, "w:ins", text, author, rev_id)


def add_del(par, text, author, rev_id):
    _revision(par, "w:del", text, author, rev_id, deleted=True)


def add_ppr_change(par, author, rev_id):
    ppr = par._p.get_or_add_pPr()
    change = OxmlElement("w:pPrChange")
    change.set(qn("w:id"), str(rev_id))
    change.set(qn("w:author"), author)
    change.set(qn("w:date"), DATE)
    change.append(OxmlElement("w:pPr"))
    ppr.append(change)


def _track_basic(d, name):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Status: ")
    add_ins(p, "new", "Ana", 1)
    add_del(p, "old", "Ben", 2)
    doc.save(str(d / name))


# ---------------------------------------------------------------- fixtures
# Each fixture writes its input files into the task folder d.

def fx_r1(d):
    doc = Document()
    doc.add_paragraph("Alpha")
    _table(doc, [["a", "b"], ["c", "d"]])
    doc.add_paragraph("Beta")
    doc.save(str(d / "r1.docx"))


def fx_c1(d):
    _write_text(d / "c1.txt", "# Title\n\nHello world\n- one\n- two\n")


def fx_c2(d):
    _write_text(d / "c2.txt",
                "## Scope\n1. first item\n1. second item\nPlain line with café\n\n- bullet\n")


def fx_c3(d):
    _write_text(d / "c3.txt", "\n\n\n")


def fx_p1(d):
    doc = Document()
    doc.add_paragraph("ACME signed the contract.")
    _table(doc, [["ACME"]])
    sec = doc.sections[0]
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "ACME Ltd letterhead"
    doc.save(str(d / "p1.docx"))


def fx_p2(d):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Contract with ")
    p.add_run("AC").bold = True
    p.add_run("ME").italic = True
    p.add_run(" Ltd signed.")
    doc.save(str(d / "p2.docx"))


def fx_p3(d):
    doc = Document()
    doc.add_paragraph("Alpha and Beta")
    doc.save(str(d / "p3.docx"))
    _write_text(d / "p3_pairs.json",
                '{\n  "Alpha": "Beta",\n  "Beta": "Gamma",\n  "Nothing": "X"\n}\n')


def fx_p4(d):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("ACME A")
    add_ins(p, "ACME B", "Ana", 1)
    add_del(p, "ACME C", "Ben", 2)
    doc.save(str(d / "p4.docx"))


def fx_f1(d):
    doc = Document()
    doc.add_paragraph("Dear {{NAME}}, due {{DATE}}.")
    _table(doc, [["Total: {{TOTAL}}"]])
    sec = doc.sections[0]
    sec.footer.is_linked_to_previous = False
    sec.footer.paragraphs[0].text = "Ref {{NAME}}"
    doc.save(str(d / "f1.docx"))
    _write_text(d / "f1.json", '{"NAME": "Ana", "DATE": "1 Nov", "TOTAL": "$40"}')


def fx_f2(d):
    doc = Document()
    doc.add_paragraph("Dear {{NAME}}, due {{DATE}}.")
    doc.save(str(d / "f2.docx"))
    _write_text(d / "f2.json", '{"NAME": "Ana"}')


def fx_f3(d):
    doc = Document()
    doc.add_paragraph("{{NAME}} {{COUNT}} {{PAID}} {{AMOUNT}}")
    doc.save(str(d / "f3.docx"))
    _write_text(d / "f3.json", '{"NAME": "Ana", "COUNT": 42, "PAID": true, "AMOUNT": 12.5}')


def fx_f4(d):
    doc = Document()
    doc.add_paragraph("Dear {{NAME}}, due {{DATE}}.")
    doc.save(str(d / "f4.docx"))
    _write_text(d / "f4.json", '{"NAME": "Ana"}')


def fx_t1(d):
    _track_basic(d, "t1.docx")


def fx_t2(d):
    _track_basic(d, "t2.docx")


def fx_t3(d):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Start ")
    add_ins(p, "new", "Ana", 1)
    add_del(p, "old", "Ben", 2)
    p.add_run(" end")
    doc.save(str(d / "t3.docx"))


def fx_t4(d):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Clause ")
    add_ins(p, "4", "Ana", 1)
    add_ppr_change(p, "Ben", 2)
    doc.save(str(d / "t4.docx"))


def fx_m1(d):
    doc = Document()
    doc.add_paragraph("The contract is signed by ACME on 1 Oct.")
    doc.save(str(d / "m1.docx"))


def fx_m2(d):
    doc = Document()
    doc.add_paragraph("ACME one; ACME two.")
    doc.save(str(d / "m2.docx"))


def fx_m3(d):
    doc = Document()
    p1 = doc.add_paragraph()
    r1 = p1.add_run("Payment due")
    doc.add_comment(runs=r1, text="Check the date", author="Ana", initials="AN")
    p2 = doc.add_paragraph("Total is ")
    r2 = p2.add_run("1,000 USD")
    doc.add_comment(runs=r2, text="Confirm the amount", author="Ben", initials="BE")
    doc.save(str(d / "m3.docx"))


def fx_v1(d):
    doc = Document()
    doc.add_paragraph("Hello")
    doc.save(str(d / "v1.docx"))


def fx_v2(d):
    doc = Document()
    doc.add_paragraph("Hello")
    full = _docx_bytes(doc)
    assert len(full) > 1000, "fixture too small to truncate"
    (d / "v2.docx").write_bytes(full[:1000])


def fx_s1(d):
    doc = Document()
    doc.add_heading("Spec", level=1)
    doc.add_heading("Scope", level=2)
    doc.add_paragraph("Para one")
    doc.add_paragraph("Para two")
    _table(doc, [["a", "b"], ["c", "d"]])
    doc.save(str(d / "s1.docx"))


def fx_pdf(d):
    doc = Document()
    doc.add_paragraph("Hello PDF")
    doc.save(str(d / "pdf.docx"))


def fx_x1(d):
    (d / "x1.doc").write_bytes(legacy_doc_bytes())


def fx_x2(d):
    (d / "x2.docx").write_bytes(encrypted_docx_bytes())


def fx_x3(d):
    doc = Document()
    doc.add_paragraph("Macro test")
    (d / "x3.docm").write_bytes(_docx_bytes(doc))


def fx_x5(d):
    doc = Document()
    doc.add_paragraph("ACME")
    doc.save(str(d / "x5.docx"))
    (d / "x5.replace.docx").write_bytes(b"EXISTING OUTPUT, NOT A DOCX")


def fx_x6(d):
    doc = Document()
    doc.add_paragraph("ACME")
    doc.save(str(d / "x6.docx"))


def fx_x7(d):
    doc = Document()
    doc.add_paragraph("Hello")
    doc.save(str(d / "x7.docx"))


# ---------------------------------------------------------------- tasks

def _task(tid, category, title, request, fixture, argv, expect):
    return {
        "id": tid,
        "category": category,
        "title": title,
        "request": request,
        "fixture": fixture,
        "argv": argv,
        "expect": expect,
    }


def build_tasks():
    """Return the 30 benchmark tasks. argv uses {dir} for the task folder."""
    return [
        # ---- read
        _task("R1", "read", "Read paragraphs and tables in order",
              "Show me everything in this file in order, tables included.",
              fx_r1, ["read", "{dir}/r1.docx"],
              {"exit": 0,
               "stdout_equals": "Alpha\na\tb\nc\td\n\nBeta",
               "new_files": []}),

        # ---- create
        _task("C1", "create", "Outline with title, body and bullets",
              "Turn this outline into a Word file.",
              fx_c1, ["create", "{dir}/c1.txt", "--out", "{dir}/c1_out.docx", "--json"],
              {"exit": 0,
               "json": {"paragraphs": 4},
               "new_files": ["c1_out.docx"],
               "output": {"file": "c1_out.docx",
                          "body_paragraphs": [["Heading 1", "Title"],
                                              ["Normal", "Hello world"],
                                              ["List Bullet", "one"],
                                              ["List Bullet", "two"]]}}),
        _task("C2", "create", "Heading 2, numbered list, UTF-8 text",
              "Make a memo from this outline; the numbered items must stay numbered.",
              fx_c2, ["create", "{dir}/c2.txt", "--out", "{dir}/c2_out.docx", "--json"],
              {"exit": 0,
               "json": {"paragraphs": 5},
               "new_files": ["c2_out.docx"],
               "output": {"file": "c2_out.docx",
                          "body_paragraphs": [["Heading 2", "Scope"],
                                              ["List Number", "first item"],
                                              ["List Number", "second item"],
                                              ["Normal", "Plain line with café"],
                                              ["List Bullet", "bullet"]]}}),
        _task("C3", "create", "Outline with no text lines",
              "Create a Word file from this (empty) outline.",
              fx_c3, ["create", "{dir}/c3.txt", "--out", "{dir}/c3_out.docx"],
              {"exit": 1, "stderr_nonempty": True, "new_files": []}),

        # ---- replace
        _task("P1", "replace", "Replace in body, table cell and header (default output name)",
              "Change every ACME to Globex in this contract, including the table and the letterhead.",
              fx_p1, ["replace", "{dir}/p1.docx", "--find", "ACME", "--replace", "Globex", "--json"],
              {"exit": 0,
               "json": {"replacements": 3,
                        "per_pair": {"ACME": 3},
                        "skipped_in_tracked_change": 0,
                        "remaining_after": {"ACME": 0},
                        "not_searched": ["footnotes", "endnotes", "comments", "field codes"]},
               "new_files": ["p1.replace.docx"],
               "output": {"file": "p1.replace.docx",
                          "body_paragraphs": [["Normal", "Globex signed the contract."]],
                          "contains": ["Globex Ltd letterhead"],
                          "absent": ["ACME"],
                          "occurrences": {"Globex": 3}}}),
        _task("P2", "replace", "Match split across runs keeps first run's formatting",
              "Replace ACME with Globex; the name is split over formatted pieces.",
              fx_p2, ["replace", "{dir}/p2.docx", "--find", "ACME", "--replace", "Globex", "--json"],
              {"exit": 0,
               "json": {"replacements": 1, "per_pair": {"ACME": 1}, "skipped_in_tracked_change": 0},
               "new_files": ["p2.replace.docx"],
               "output": {"file": "p2.replace.docx",
                          "body_paragraphs": [["Normal", "Contract with Globex Ltd signed."]],
                          "absent": ["ACME"],
                          "run_format": {"text": "Globex", "bold": True, "italic": False}}}),
        _task("P3", "replace", "Pairs file applied in order, with a zero-match pair",
              "Run these replacements in order: Alpha to Beta, then Beta to Gamma, then Nothing to X.",
              fx_p3, ["replace", "{dir}/p3.docx", "--pairs", "{dir}/p3_pairs.json", "--json"],
              {"exit": 0,
               "json": {"replacements": 3, "per_pair": {"Alpha": 1, "Beta": 2, "Nothing": 0}},
               "new_files": ["p3.replace.docx"],
               "output": {"file": "p3.replace.docx",
                          "body_paragraphs": [["Normal", "Gamma and Gamma"]]}}),
        _task("P4", "replace", "Skip matches inside tracked insertions, ignore deletions",
              "Replace ACME with Globex but leave other people's tracked edits alone.",
              fx_p4, ["replace", "{dir}/p4.docx", "--find", "ACME", "--replace", "Globex", "--json"],
              {"exit": 0,
               "json": {"replacements": 1, "per_pair": {"ACME": 1}, "skipped_in_tracked_change": 1},
               "new_files": ["p4.replace.docx"],
               "output": {"file": "p4.replace.docx",
                          "occurrences": {"Globex": 1, "ACME": 1},
                          "revisions": {"insert": 1, "delete": 1, "delText": 1},
                          "deleted_contains": ["ACME C"]}}),

        # ---- fill
        _task("F1", "fill", "Fill placeholders in body, table and footer",
              "Fill this letter template with the client details.",
              fx_f1, ["fill", "{dir}/f1.docx", "--values", "{dir}/f1.json", "--json"],
              {"exit": 0,
               "json": {"filled": 4, "per_name": {"NAME": 2, "DATE": 1, "TOTAL": 1}, "unfilled": []},
               "new_files": ["f1.fill.docx"],
               "output": {"file": "f1.fill.docx",
                          "body_paragraphs": [["Normal", "Dear Ana, due 1 Nov."]],
                          "contains": ["Total: $40", "Ref Ana"],
                          "absent": ["{{"]}}),
        _task("F2", "fill", "Missing value stops the fill and writes nothing",
              "Fill the letter; I only have the name, not the date.",
              fx_f2, ["fill", "{dir}/f2.docx", "--values", "{dir}/f2.json"],
              {"exit": 1, "message": ["Missing values for: DATE"], "new_files": []}),
        _task("F3", "fill", "Numbers and booleans are converted as text",
              "Fill the form with the count, the paid flag and the amount.",
              fx_f3, ["fill", "{dir}/f3.docx", "--values", "{dir}/f3.json", "--json"],
              {"exit": 0,
               "json": {"filled": 4},
               "new_files": ["f3.fill.docx"],
               "output": {"file": "f3.fill.docx",
                          "body_paragraphs": [["Normal", "Ana 42 True 12.5"]]}}),
        _task("F4", "fill", "Allow unfilled placeholders and report them",
              "Fill what you can in the letter and leave the date as a placeholder.",
              fx_f4, ["fill", "{dir}/f4.docx", "--values", "{dir}/f4.json", "--allow-unfilled", "--json"],
              {"exit": 0,
               "json": {"filled": 1, "unfilled": ["DATE"]},
               "new_files": ["f4.fill.docx"],
               "output": {"file": "f4.fill.docx",
                          "body_paragraphs": [["Normal", "Dear Ana, due {{DATE}}."]]}}),

        # ---- track-changes
        _task("T1", "track-changes", "List an insertion and a deletion",
              "Show me the tracked changes in this file and who made them.",
              fx_t1, ["track-show", "{dir}/t1.docx", "--json"],
              {"exit": 0,
               "json": {"revisions.#": 2,
                        "revisions.0.kind": "insert",
                        "revisions.0.author": "Ana",
                        "revisions.0.text": "new",
                        "revisions.0.location": "body",
                        "revisions.1.kind": "delete",
                        "revisions.1.author": "Ben",
                        "revisions.1.text": "old",
                        "revisions.1.location": "body"},
               "new_files": []}),
        _task("T2", "track-changes", "Accept all tracked changes",
              "Accept all the tracked changes and give me a clean copy.",
              fx_t2, ["track-resolve", "{dir}/t2.docx", "--mode", "accept", "--all", "--json"],
              {"exit": 0,
               "json": {"resolved": 2, "remaining": 0},
               "new_files": ["t2.track-resolve.docx"],
               "output": {"file": "t2.track-resolve.docx",
                          "body_paragraphs": [["Normal", "Status: new"]],
                          "revisions": {"insert": 0, "delete": 0, "delText": 0}}}),
        _task("T3", "track-changes", "Reject all tracked changes",
              "Reject all the tracked changes in this document.",
              fx_t3, ["track-resolve", "{dir}/t3.docx", "--mode", "reject", "--all", "--json"],
              {"exit": 0,
               "json": {"resolved": 2, "remaining": 0},
               "new_files": ["t3.track-resolve.docx"],
               "output": {"file": "t3.track-resolve.docx",
                          "body_paragraphs": [["Normal", "Start old end"]],
                          "revisions": {"insert": 0, "delete": 0, "delText": 0}}}),
        _task("T4", "track-changes", "Refuse to accept a formatting change",
              "Accept all the changes, including the formatting change.",
              fx_t4, ["track-resolve", "{dir}/t4.docx", "--mode", "accept", "--all"],
              {"exit": 2,
               "message": ["cannot be resolved by this skill"],
               "new_files": []}),

        # ---- comments
        _task("M1", "comments", "Add a comment anchored to exact text",
              "Add a comment on the word 'signed' saying 'Check date', author Test Bot.",
              fx_m1, ["comments-add", "{dir}/m1.docx", "--anchor", "signed", "--text", "Check date",
                      "--author", "Test Bot", "--json"],
              {"exit": 0,
               "json": {"anchor_text": "signed", "text": "Check date", "author": "Test Bot"},
               "new_files": ["m1.comments-add.docx"],
               "output": {"file": "m1.comments-add.docx",
                          "comments": [{"author": "Test Bot", "text": "Check date", "anchor": "signed"}],
                          "body_paragraphs": [["Normal", "The contract is signed by ACME on 1 Oct."]]}}),
        _task("M2", "comments", "Ambiguous anchor without --occurrence",
              "Add a comment on ACME, Test Bot.",
              fx_m2, ["comments-add", "{dir}/m2.docx", "--anchor", "ACME", "--text", "Which one?",
                      "--author", "Test Bot"],
              {"exit": 1, "message": ["Anchor text occurs 2 times"], "new_files": []}),
        _task("M3", "comments", "List comments with anchors and authors",
              "List all the review comments in this file.",
              fx_m3, ["comments-list", "{dir}/m3.docx", "--json"],
              {"exit": 0,
               "json": {"comments.#": 2,
                        "comments.0.author": "Ana",
                        "comments.0.initials": "AN",
                        "comments.0.text": "Check the date",
                        "comments.0.anchor_text": "Payment due",
                        "comments.1.author": "Ben",
                        "comments.1.initials": "BE",
                        "comments.1.text": "Confirm the amount",
                        "comments.1.anchor_text": "1,000 USD"},
               "new_files": []}),

        # ---- validate and corrupt input
        _task("V1", "validate", "Validate a good file",
              "Check that this Word file opens and is not damaged.",
              fx_v1, ["validate", "{dir}/v1.docx", "--json"],
              {"exit": 0,
               "json": {"verdict": "valid"},
               "checks_all_passed": True,
               "new_files": []}),
        _task("V2", "corrupt-file", "Validate a truncated file",
              "Check whether this Word file is damaged (it was cut off during a download).",
              fx_v2, ["validate", "{dir}/v2.docx", "--json"],
              {"exit": 1,
               "json": {"verdict": "invalid", "checks.0.passed": False},
               "new_files": []}),

        # ---- summary
        _task("S1", "summary", "Structure summary of a small report",
              "Describe the structure of this report: headings, paragraph count and tables.",
              fx_s1, ["summary", "{dir}/s1.docx", "--json"],
              {"exit": 0,
               "json": {"headings": [{"level": 1, "text": "Spec"}, {"level": 2, "text": "Scope"}],
                        "paragraph_count": 4,
                        "table_count": 1,
                        "table_sizes": [[2, 2]],
                        "section_count": 1,
                        "has_comments": False,
                        "has_tracked_changes": False,
                        "styles_used": {"Heading 1": 1, "Heading 2": 1, "Normal": 2}},
               "new_files": []}),

        # ---- to-pdf (expectation depends on the converter found: Word, LibreOffice or none)
        _task("PDF1", "to-pdf", "Convert to PDF with Word or LibreOffice, or refuse when neither is installed",
              "Export this letter as a PDF.",
              fx_pdf, ["to-pdf", "{dir}/pdf.docx", "--out", "{dir}/pdf.pdf", "--timeout", "300"],
              {"variants": {
                  "word": {"exit": 0, "output": {"file": "pdf.pdf", "pdf_pages": 1}},
                  "libreoffice": {"exit": 0, "output": {"file": "pdf.pdf", "pdf_pages": 1}},
                  "none": {"exit": 2,
                           "message": ["LibreOffice is not installed"],
                           "new_files": []},
              }}),

        # ---- out-of-scope inputs and requests
        _task("X1", "out-of-scope", "Legacy .doc file is refused",
              "Read this old Word 97 file.",
              fx_x1, ["read", "{dir}/x1.doc"],
              {"exit": 2, "message": ["Legacy .doc is not supported"], "new_files": []}),
        _task("X2", "out-of-scope", "Password-protected file is refused",
              "Summarize this file, it is the password-protected contract.",
              fx_x2, ["summary", "{dir}/x2.docx", "--json"],
              {"exit": 2, "message": ["password-protected"], "new_files": []}),
        _task("X3", "out-of-scope", "Macro-enabled .docm is refused",
              "Read this macro workbook-style Word file.",
              fx_x3, ["read", "{dir}/x3.docm"],
              {"exit": 2, "message": ["Macro-enabled .docm files are not supported"], "new_files": []}),
        _task("X4", "missing-file", "Missing input file",
              "Read the contract in this folder.",
              None, ["read", "{dir}/x4_missing.docx"],
              {"exit": 1, "message": ["Input file not found"], "new_files": []}),
        _task("X5", "safety", "Existing output is not replaced without --force",
              "Replace ACME with Globex and save it as the file that is already there.",
              fx_x5, ["replace", "{dir}/x5.docx", "--find", "ACME", "--replace", "Globex"],
              {"exit": 1, "message": ["Output already exists"], "new_files": []}),
        _task("X6", "safety", "Output equal to input is refused without --in-place",
              "Replace ACME with Globex in the same file.",
              fx_x6, ["replace", "{dir}/x6.docx", "--find", "ACME", "--replace", "Globex",
                      "--out", "{dir}/x6.docx"],
              {"exit": 1, "message": ["Refusing to overwrite the input"], "new_files": []}),
        _task("X7", "out-of-scope", "Request for a command the tool does not have",
              "Insert the company logo picture into this document.",
              fx_x7, ["insert-image", "{dir}/x7.docx"],
              {"exit": 1, "stderr_nonempty": True, "new_files": []}),
    ]


# ---------------------------------------------------------------- fact readers

_PART_RE = re.compile(r"word/(document|header\d*|footer\d*)\.xml$")


def _para_text(el):
    """Visible text of a paragraph: w:t inside it, tracked insertions included, deletions excluded."""
    return "".join(t.text or "" for t in el.iter(T_TAG))


def _flag(rpr, name):
    if rpr is None:
        return False
    el = rpr.find(f"{{{W}}}{name}")
    if el is None:
        return False
    return el.get(f"{{{W}}}val") not in ("0", "false", "off")


def _anchors(root):
    anchors, active = {}, []
    for el in root.iter():
        if el.tag == f"{{{W}}}commentRangeStart":
            cid = el.get(f"{{{W}}}id")
            anchors[cid] = ""
            active.append(cid)
        elif el.tag == f"{{{W}}}commentRangeEnd":
            cid = el.get(f"{{{W}}}id")
            if cid in active:
                active.remove(cid)
        elif el.tag == T_TAG:
            for cid in active:
                anchors[cid] += el.text or ""
    return anchors


def collect_facts(path):
    """Read a .docx independently of the tool and return its facts."""
    path = Path(path)
    doc = Document(str(path))
    body = [[Paragraph(el, doc).style.name, _para_text(el)]
            for el in doc.element.body.iterchildren(P_TAG)]
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        roots = {n: etree.fromstring(z.read(n)) for n in names if _PART_RE.match(n)}
        comments_root = etree.fromstring(z.read("word/comments.xml")) if "word/comments.xml" in names else None
    main = roots["word/document.xml"]
    visible = "\n".join(_para_text(p) for name in sorted(roots) for p in roots[name].iter(P_TAG))
    deleted = "\n".join(t.text or "" for t in main.iter(DT_TAG))
    revisions = {
        "insert": sum(1 for _ in main.iter(f"{{{W}}}ins")),
        "delete": sum(1 for _ in main.iter(f"{{{W}}}del")),
        "delText": sum(1 for _ in main.iter(DT_TAG)),
        "pPrChange": sum(1 for _ in main.iter(f"{{{W}}}pPrChange")),
        "rPrChange": sum(1 for _ in main.iter(f"{{{W}}}rPrChange")),
    }
    runs = []
    for r in main.iter(R_TAG):
        rpr = r.find(f"{{{W}}}rPr")
        runs.append({"text": "".join(t.text or "" for t in r.iter(T_TAG)),
                     "bold": _flag(rpr, "b"), "italic": _flag(rpr, "i")})
    comments = []
    if comments_root is not None:
        anchors = _anchors(main)
        for c in comments_root.iter(f"{{{W}}}comment"):
            cid = c.get(f"{{{W}}}id")
            comments.append({
                "author": c.get(f"{{{W}}}author"),
                "initials": c.get(f"{{{W}}}initials"),
                "text": "\n".join(_para_text(p) for p in c.iter(P_TAG)),
                "anchor": anchors.get(cid),
            })
    return {"body": body, "visible": visible, "deleted": deleted,
            "revisions": revisions, "runs": runs, "comments": comments}


def pdf_pages(path):
    from pypdf import PdfReader
    return len(PdfReader(str(path)).pages)
