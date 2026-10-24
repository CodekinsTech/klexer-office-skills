#!/usr/bin/env python3
"""pdf_tool.py - one command-line entry point for everyday PDF jobs.

Usage:  python pdf_tool.py COMMAND [options]      (python pdf_tool.py --help)

Exit codes: 0 done (and verified where the command writes a file),
            1 bad input, bad option, or a failed check,
            2 a password is needed, or the job is out of scope.
Errors go to stderr. Read commands print JSON with --json, text otherwise.
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import math
import os
import re
import sys
import tempfile
from collections import Counter
from pathlib import Path

from pypdf import PasswordType, PdfReader, PdfWriter
from pypdf.errors import PdfReadError, PdfStreamError
from pypdf.generic import NameObject, NumberObject

EXIT_OK = 0
EXIT_FAIL = 1
EXIT_BLOCKED = 2

DEFAULT_PASSWORD_ENV = "PDF_PASSWORD"
ALLOWED_ANGLES = ("90", "180", "270")
POINTS_PER_MM = 72 / 25.4
NO_TEXT_LAYER = "No text layer found. Scanned PDFs need OCR, which this skill does not do."
STAMP_OR_TEXT = "Give exactly one of --stamp or --text."
LOCKED_MESSAGE = "The file is open in another program. Close it and try again."
IN_PLACE_ENCRYPTED = "--in-place is refused for a password-protected PDF. Write to a new file with --out."

# AcroForm field flags (PDF 1.7 table 221 and 226), as bit values.
FLAG_RADIO = 1 << 15
FLAG_PUSHBUTTON = 1 << 16
FLAG_COMBO = 1 << 17


class ToolError(Exception):
    """A refusal or failure. `message` is printed to stderr exactly as given."""

    def __init__(self, message: str, code: int = EXIT_FAIL, kind: str = ""):
        super().__init__(message)
        self.message = message
        self.code = code
        self.kind = kind


class VerifyError(Exception):
    """Raised by a verification step when a written file does not match the request."""


def corrupt_message(path) -> str:
    return f"File is not a readable PDF (corrupt or truncated): {path}"


def password_message(name: str) -> str:
    return (f"This PDF is password-protected. Set the environment variable {name} "
            f"and pass --password-env {name}.")


# ---------------------------------------------------------------- argument parsing

class _Parser(argparse.ArgumentParser):
    """Argparse exits 1 (not 2) on usage errors, and the message names the option."""

    def error(self, message: str):  # type: ignore[override]
        sys.stderr.write(f"{self.prog}: error: {message}\n")
        sys.exit(EXIT_FAIL)


def _add_read_common(p: argparse.ArgumentParser) -> None:
    p.add_argument("--password-env", metavar="NAME",
                   help="name of the environment variable that holds the PDF password")
    p.add_argument("--json", action="store_true", help="print one JSON object")
    p.add_argument("--out", metavar="FILE", help="write the same result to a new file")
    p.add_argument("--force", action="store_true", help="replace --out if it exists")


def _add_write_common(p: argparse.ArgumentParser, in_place: bool = False) -> None:
    if in_place:
        # Exactly one of --out or --in-place. argparse names both if neither or both are given.
        group = p.add_mutually_exclusive_group(required=True)
        group.add_argument("--out", metavar="FILE", help="new PDF to write")
        group.add_argument("--in-place", action="store_true",
                           help="replace the input with the result, after the result is checked")
    else:
        p.add_argument("--out", metavar="FILE", required=True, help="new PDF to write")
    p.add_argument("--force", action="store_true", help="replace --out if it exists")
    p.add_argument("--json", action="store_true", help="print one JSON object")


def build_parser() -> argparse.ArgumentParser:
    parser = _Parser(prog="pdf_tool.py", description="Read, check, fill and edit PDF files.")
    sub = parser.add_subparsers(dest="command", metavar="COMMAND", parser_class=_Parser)

    p = sub.add_parser("text", help="extract text page by page")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--pages", metavar="LIST", help="e.g. 1-3,5,8- (default all)")
    _add_read_common(p)
    p.set_defaults(handler=cmd_text)

    p = sub.add_parser("tables", help="extract tables from pages")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--pages", metavar="LIST", help="e.g. 1-3,5,8- (default all)")
    _add_read_common(p)
    p.set_defaults(handler=cmd_tables)

    p = sub.add_parser("fields-list", help="list AcroForm fields")
    p.add_argument("input_file", metavar="INPUT.pdf")
    _add_read_common(p)
    p.set_defaults(handler=cmd_fields_list)

    p = sub.add_parser("fields-fill", help="fill AcroForm fields from a JSON file")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--values", metavar="FILE.json", required=True,
                   help="JSON object of field name to value")
    p.add_argument("--password-env", metavar="NAME")
    _add_write_common(p, in_place=True)
    p.set_defaults(handler=cmd_fields_fill)

    p = sub.add_parser("merge", help="join PDFs in the order given")
    p.add_argument("inputs", metavar="INPUT.pdf", nargs="*")
    p.add_argument("--password-env", metavar="NAME")
    _add_write_common(p)
    p.set_defaults(handler=cmd_merge)

    p = sub.add_parser("split", help="split a PDF into parts in a folder")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--every", metavar="N", help="pages per part")
    p.add_argument("--ranges", metavar="LIST", help='e.g. "1-3,4-6"; must cover every page once')
    p.add_argument("--out-dir", metavar="DIR", required=True,
                   help="folder to create, or an existing empty folder")
    p.add_argument("--password-env", metavar="NAME")
    p.add_argument("--json", action="store_true", help="print one JSON object")
    p.set_defaults(handler=cmd_split)

    p = sub.add_parser("rotate", help="rotate pages clockwise")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--pages", metavar="LIST", help="default all pages")
    p.add_argument("--angle", metavar="DEGREES", required=True, help="90, 180 or 270")
    p.add_argument("--password-env", metavar="NAME")
    _add_write_common(p, in_place=True)
    p.set_defaults(handler=cmd_rotate)

    p = sub.add_parser("reorder", help="reorder every page with a permutation")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--order", metavar="LIST", required=True, help="e.g. 3,1,2")
    p.add_argument("--password-env", metavar="NAME")
    _add_write_common(p, in_place=True)
    p.set_defaults(handler=cmd_reorder)

    p = sub.add_parser("watermark", help="add a stamp PDF or diagonal text")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--stamp", metavar="FILE.pdf", help="first page is drawn over each page")
    p.add_argument("--text", metavar="TEXT", help="diagonal text drawn over each page")
    p.add_argument("--pages", metavar="LIST", help="default all pages")
    p.add_argument("--password-env", metavar="NAME")
    _add_write_common(p, in_place=True)
    p.set_defaults(handler=cmd_watermark)

    p = sub.add_parser("create", help="make a PDF from a UTF-8 text file")
    p.add_argument("input_file", metavar="INPUT.txt")
    p.add_argument("--font-size", metavar="N", default="11")
    p.add_argument("--page-size", metavar="A4|Letter", default="A4", choices=["A4", "Letter"])
    p.add_argument("--margin", metavar="MM", default="20")
    _add_write_common(p)
    p.set_defaults(handler=cmd_create)

    p = sub.add_parser("check", help="check that a PDF opens and reads cleanly")
    p.add_argument("input_file", metavar="INPUT.pdf")
    _add_read_common(p)
    p.set_defaults(handler=cmd_check)

    p = sub.add_parser("check-text", help="check that expected phrases appear in a PDF")
    p.add_argument("input_file", metavar="INPUT.pdf")
    p.add_argument("--expect", metavar="FILE.txt", required=True,
                   help="one phrase per line")
    _add_read_common(p)
    p.set_defaults(handler=cmd_check_text)

    return parser


# ---------------------------------------------------------------- shared helpers

def eprint(message: str) -> None:
    sys.stderr.write(message + "\n")


def _norm(text: str) -> str:
    """Collapse every run of whitespace to one space and strip the ends."""
    return " ".join(text.split())


def _json_text(result) -> str:
    return json.dumps(result, ensure_ascii=False, indent=2)


def parse_page_list(text, total: int) -> list[int]:
    """Return sorted, unique 1-based pages for a list such as '1-3,5,8-'. None means all."""
    if text is None:
        return list(range(1, total + 1))
    pages: set[int] = set()
    for start, end in _page_items(text, total):
        pages.update(range(start, end + 1))
    return sorted(pages)


def _page_items(text: str, total: int) -> list[tuple[int, int]]:
    items = []
    for part in text.split(","):
        match = re.fullmatch(r"\s*(\d+)(?:-(\d*))?\s*", part)
        bad = ToolError(f"Page list is invalid or outside 1..{total}: {text}", EXIT_FAIL)
        if not match:
            raise bad
        start = int(match.group(1))
        if match.group(2) is None:
            end = start
        elif match.group(2) == "":
            end = total
        else:
            end = int(match.group(2))
        if start < 1 or end > total or start > end:
            raise bad
        items.append((start, end))
    return items


def parse_ranges(text: str, total: int) -> list[tuple[int, int]]:
    """Ranges for split, in the order given. Every page must be covered exactly once."""
    items = _page_items(text, total)
    counts: Counter = Counter()
    for start, end in items:
        counts.update(range(start, end + 1))
    for page in range(1, total + 1):
        if counts[page] == 0:
            raise ToolError(f"Ranges must cover each page once. Page {page} is not covered.")
        if counts[page] > 1:
            raise ToolError(f"Ranges must cover each page once. Page {page} is covered more than once.")
    return items


def prepare_output(out: Path, force: bool, sources) -> None:
    """Refuse an input as output, a folder, or an existing file without --force."""
    for src in sources:
        if os.path.normcase(str(Path(src).resolve())) == os.path.normcase(str(out.resolve())):
            raise ToolError("Refusing to overwrite an input. Write to a new file.", EXIT_FAIL)
    if out.is_dir():
        raise ToolError(f"Output path is a folder: {out}", EXIT_FAIL)
    if out.exists() and not force:
        raise ToolError(f"Output already exists: {out}. Choose another --out or pass --force.", EXIT_FAIL)


def resolve_target(args, src: Path, reader, *others) -> Path:
    """The file a write command publishes: the --out file, or the input itself with --in-place.

    --in-place is refused for a password-protected input, so decryption can never happen by
    accident. No backup is kept: the spec does not ask for one, and the result is checked before
    it replaces the input.
    """
    if not args.in_place:
        out = Path(args.out)
        prepare_output(out, args.force, [src, *others])
        return out
    if reader.is_encrypted:
        raise ToolError(IN_PLACE_ENCRYPTED, EXIT_FAIL)
    return src


def open_pdf(path: Path, password_env, strict: bool = False):
    """Open a PDF, applying the password rules. Returns (reader, password_used)."""
    if not path.is_file():
        raise ToolError(f"Input file not found: {path}", EXIT_FAIL)
    try:
        data = path.read_bytes()
    except PermissionError:  # another program holds the file with a lock
        raise ToolError(LOCKED_MESSAGE, EXIT_FAIL) from None
    try:
        reader = PdfReader(io.BytesIO(data), strict=strict)
        encrypted = reader.is_encrypted
    except Exception:  # any parse failure means the file is not a usable PDF
        raise ToolError(corrupt_message(path), EXIT_FAIL, kind="corrupt") from None
    used = ""
    if encrypted:
        name = password_env or DEFAULT_PASSWORD_ENV
        supplied = os.environ.get(password_env, "") if password_env else ""
        if supplied:
            if int(reader.decrypt(supplied)) == int(PasswordType.NOT_DECRYPTED):
                raise ToolError(f"The password in {name} did not open this PDF.", EXIT_FAIL)
            used = supplied
        else:
            # No password supplied: only a PDF whose user password is empty can be opened.
            if int(reader.decrypt("")) == int(PasswordType.NOT_DECRYPTED):
                raise ToolError(password_message(name), EXIT_BLOCKED)
    return reader, used


def _check_out(args, *sources) -> None:
    if getattr(args, "out", None):
        prepare_output(Path(args.out), args.force, sources)


def _emit(args, rendered: str) -> None:
    """Print a read result, or write it to --out with the same re-open check."""
    if args.out:
        data = (rendered + "\n").encode("utf-8")
        _commit(Path(args.out), lambda fh: fh.write(data), lambda p: _expect_bytes(p, data))
    else:
        sys.stdout.write(rendered + "\n")


def _expect_bytes(path: Path, data: bytes) -> None:
    if path.read_bytes() != data:
        raise VerifyError("the written file does not match the result")


def _report_write(args, result: dict, line: str) -> int:
    if args.json:
        sys.stdout.write(_json_text(result) + "\n")
    else:
        sys.stdout.write(line + "\n")
    return EXIT_OK


def _stage(folder: Path, produce) -> Path:
    """Write content to a temp file inside `folder` (same volume, so os.replace is atomic)."""
    folder.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=".pdf-tool-", suffix=".tmp", dir=str(folder))
    tmp = Path(name)
    try:
        with os.fdopen(fd, "wb") as fh:
            produce(fh)
    except BaseException:
        _discard(tmp)
        raise
    return tmp


def _verify_path(path: Path, check) -> None:
    try:
        check(path)
    except VerifyError as exc:
        raise ToolError(f"Output failed verification and was not kept: {exc}", EXIT_FAIL) from None
    except Exception as exc:
        raise ToolError(f"Output failed verification and was not kept: {type(exc).__name__}",
                        EXIT_FAIL) from None


def _publish(tmp: Path, final: Path) -> None:
    try:
        os.replace(tmp, final)
    except PermissionError:
        raise ToolError(LOCKED_MESSAGE, EXIT_FAIL) from None


def _discard(path: Path) -> None:
    try:
        path.unlink()
    except OSError:
        pass


def _commit(out: Path, produce, check) -> None:
    """Stage, re-open and check, then move into place. A failed check keeps nothing."""
    tmp = _stage(out.parent, produce)
    try:
        _verify_path(tmp, check)
        _publish(tmp, out)
    finally:
        _discard(tmp)


def _pdf_check(fn):
    """Adapt a function that takes a PdfReader into a path check."""
    def check(path: Path) -> None:
        fn(PdfReader(io.BytesIO(path.read_bytes())))
    return check


def _expect_count(reader, expected: int) -> None:
    if len(reader.pages) != expected:
        raise VerifyError(f"it has {len(reader.pages)} pages, expected {expected}")


def _is_subsequence(needle: list[str], hay: list[str]) -> bool:
    it = iter(hay)
    return all(any(token == h for h in it) for token in needle)


# ---------------------------------------------------------------- read commands

def cmd_text(args) -> int:
    src = Path(args.input_file)
    _check_out(args, src)
    reader, _ = open_pdf(src, args.password_env)
    total = len(reader.pages)
    selected = parse_page_list(args.pages, total)
    pages, no_text = [], []
    for number in selected:
        text = reader.pages[number - 1].extract_text() or ""
        if not text.strip():
            text = ""
            no_text.append(number)
        pages.append({"page": number, "text": text})
    if selected and len(no_text) == len(selected):
        raise ToolError(NO_TEXT_LAYER, EXIT_BLOCKED)
    result = {"page_count": total, "pages": pages, "no_text_pages": no_text}
    if args.json:
        rendered = _json_text(result)
    else:
        rendered = "\n".join(f"--- page {p['page']} ---\n{p['text']}" for p in pages)
    _emit(args, rendered)
    return EXIT_OK


def cmd_tables(args) -> int:
    from pdfminer.pdfdocument import PDFException
    import pdfplumber

    src = Path(args.input_file)
    _check_out(args, src)
    reader, used = open_pdf(src, args.password_env)
    total = len(reader.pages)
    selected = parse_page_list(args.pages, total)
    tables = []
    try:
        data = src.read_bytes()
        with pdfplumber.open(io.BytesIO(data), password=used if reader.is_encrypted else None) as pdf:
            for number in selected:
                found = pdf.pages[number - 1].extract_tables()
                for index, table in enumerate(found, start=1):
                    rows = [["" if cell is None else str(cell) for cell in row] for row in table]
                    tables.append({"page": number, "index": index, "rows": rows})
    except PDFException:
        raise ToolError(corrupt_message(src), EXIT_FAIL) from None
    result = {"tables": tables}
    if args.json:
        rendered = _json_text(result)
    elif not tables:
        rendered = "No tables found."
    else:
        blocks = []
        for t in tables:
            lines = [f"--- page {t['page']} table {t['index']} ---"]
            lines += ["\t".join(cell.replace("\n", " ") for cell in row) for row in t["rows"]]
            blocks.append("\n".join(lines))
        rendered = "\n".join(blocks)
    _emit(args, rendered)
    return EXIT_OK


FIELD_KINDS = ("text", "checkbox", "dropdown", "radio", "other")


def _field_kind(field) -> str:
    ft = str(field.get("/FT", ""))
    flags = int(field.get("/Ff", 0) or 0)
    if ft == "/Tx":
        return "text"
    if ft == "/Btn":
        if flags & FLAG_RADIO:
            return "radio"
        if flags & FLAG_PUSHBUTTON:
            return "other"
        return "checkbox"
    if ft == "/Ch":
        return "dropdown" if flags & FLAG_COMBO else "other"
    return "other"


def _qualified_name(node) -> str:
    parts = []
    depth = 0
    while node is not None and depth < 64:
        if "/T" in node:
            parts.append(str(node["/T"]))
        parent = node.get("/Parent")
        node = parent.get_object() if parent is not None else None
        depth += 1
    return ".".join(reversed(parts))


def _widgets_by_name(reader) -> dict:
    """Map each field's full name to the (page number, widget annotation) pairs."""
    found: dict = {}
    for index, page in enumerate(reader.pages, start=1):
        for ref in page.get("/Annots") or []:
            annot = ref.get_object()
            if str(annot.get("/Subtype", "")) != "/Widget":
                continue
            name = _qualified_name(annot)
            if name:
                found.setdefault(name, []).append((index, annot))
    return found


def _appearance_states(annots) -> list[str]:
    states: list[str] = []
    for _, annot in annots:
        ap = annot.get("/AP")
        if ap is None:
            continue
        normal = ap.get_object().get("/N")
        if normal is None:
            continue
        normal = normal.get_object()
        if not hasattr(normal, "keys"):
            continue
        for key in normal.keys():
            if str(key) not in states:
                states.append(str(key))
    return states


def _field_rows(reader) -> list[dict]:
    """One row per form field. Keys starting with '_' are internal and not printed."""
    fields = reader.get_fields() or {}
    if not fields:
        return []
    widgets = _widgets_by_name(reader)
    rows = []
    for name, field in fields.items():
        kind = _field_kind(field)
        annots = widgets.get(name, [])
        page = annots[0][0] if annots else None
        raw = field.get("/V")
        raw_text = "" if raw is None else str(raw)
        states = _appearance_states(annots) if kind in ("checkbox", "radio") else []
        options: list[str] = []
        if kind == "radio":
            options = [s[1:] for s in states if s != "/Off"]
        elif kind == "dropdown":
            for item in field.get("/Opt") or []:
                item = item.get_object()
                options.append(str(item[0]) if isinstance(item, list) else str(item))
        if kind == "checkbox":
            value = raw_text not in ("", "/Off", "Off")
        elif kind == "radio":
            value = raw_text[1:] if raw_text.startswith("/") else raw_text
            value = "" if value == "Off" else value
        else:
            value = raw_text
        rows.append({"name": name, "type": kind, "value": value, "options": options,
                     "page": page, "_states": states})
    return rows


def _public(rows: list[dict]) -> list[dict]:
    return [{k: v for k, v in row.items() if not k.startswith("_")} for row in rows]


def cmd_fields_list(args) -> int:
    src = Path(args.input_file)
    _check_out(args, src)
    reader, _ = open_pdf(src, args.password_env)
    rows = _public(_field_rows(reader))
    result = {"fields": rows}
    if args.json:
        rendered = _json_text(result)
    elif not rows:
        rendered = "No form fields."
    else:
        lines = ["name\ttype\tvalue\toptions\tpage"]
        for r in rows:
            value = str(r["value"]).lower() if r["type"] == "checkbox" else r["value"]
            lines.append("\t".join([r["name"], r["type"], str(value),
                                    ", ".join(r["options"]),
                                    "" if r["page"] is None else str(r["page"])]))
        rendered = "\n".join(lines)
    _emit(args, rendered)
    return EXIT_OK


def cmd_check(args) -> int:
    src = Path(args.input_file)
    if not src.is_file():
        raise ToolError(f"Input file not found: {src}", EXIT_FAIL)
    _check_out(args, src)
    errors: list[str] = []
    warnings: list[str] = []
    page_count = 0
    encrypted = False
    capture = _WarningCapture()
    pypdf_logger = logging.getLogger("pypdf")
    pypdf_logger.addHandler(capture)
    try:
        reader = None
        try:
            # Strict mode: a broken cross-reference table or trailer is an error, not a repair.
            reader, _ = open_pdf(src, args.password_env, strict=True)
        except ToolError as exc:
            if exc.kind != "corrupt":
                raise
            errors.append(exc.message)
        if reader is not None:
            encrypted = bool(reader.is_encrypted)
            try:
                page_count = len(reader.pages)
            except Exception as exc:
                errors.append(f"Page tree could not be read: {type(exc).__name__}.")
                page_count = 0
            if page_count < 1 and not errors:
                errors.append("PDF has no pages.")
            any_text = False
            for number in range(1, page_count + 1):
                try:
                    page = reader.pages[number - 1]
                    contents = page.get_contents()
                    if contents is not None:
                        contents.get_data()
                    if (page.extract_text() or "").strip():
                        any_text = True
                except Exception as exc:
                    errors.append(f"Page {number} could not be read: {type(exc).__name__}.")
            if page_count and not any_text:
                warnings.append("No text layer found on any page; the file may be scanned.")
    finally:
        pypdf_logger.removeHandler(capture)
    for message in capture.messages:
        if message not in warnings and message not in errors:
            warnings.append(message)
    result = {"readable": not errors, "page_count": page_count, "encrypted": encrypted,
              "warnings": warnings, "errors": errors}
    if args.json:
        rendered = _json_text(result)
    else:
        rendered = "\n".join([
            f"readable: {'yes' if not errors else 'no'}",
            f"page_count: {page_count}",
            f"encrypted: {'yes' if encrypted else 'no'}",
            "warnings:" + ("" if not warnings else "\n  " + "\n  ".join(warnings)),
            "errors:" + ("" if not errors else "\n  " + "\n  ".join(errors)),
        ])
    _emit(args, rendered)
    for message in errors:
        eprint(message)
    return EXIT_OK if not errors else EXIT_FAIL


class _WarningCapture(logging.Handler):
    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.messages: list[str] = []

    def emit(self, record):
        self.messages.append(record.getMessage())


def _read_json_object(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        raise ToolError(f"Could not read {path}: file not found.", EXIT_FAIL) from None
    except UnicodeDecodeError:
        raise ToolError(f"Could not read {path}: not UTF-8 text.", EXIT_FAIL) from None
    except (OSError, json.JSONDecodeError) as exc:
        raise ToolError(f"Could not read {path}: {exc}", EXIT_FAIL) from None
    if not isinstance(data, dict):
        raise ToolError(f"Could not read {path}: the JSON must be an object of field names to values.",
                        EXIT_FAIL)
    return data


def _read_phrases(path: Path) -> list[str]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        raise ToolError(f"Could not read {path}: file not found.", EXIT_FAIL) from None
    except UnicodeDecodeError:
        raise ToolError(f"Could not read {path}: not UTF-8 text.", EXIT_FAIL) from None
    except OSError as exc:
        raise ToolError(f"Could not read {path}: {exc}", EXIT_FAIL) from None
    phrases = [p for p in (_norm(line) for line in text.splitlines()) if p]
    if not phrases:
        raise ToolError(f"Could not read {path}: it has no phrases.", EXIT_FAIL)
    return phrases


def cmd_check_text(args) -> int:
    src = Path(args.input_file)
    _check_out(args, src)
    phrases = _read_phrases(Path(args.expect))
    reader, _ = open_pdf(src, args.password_env)
    texts = [_norm(page.extract_text() or "") for page in reader.pages]
    if not any(texts):
        raise ToolError(NO_TEXT_LAYER, EXIT_BLOCKED)
    found, missing = [], []
    for phrase in phrases:
        hit = next((i + 1 for i, text in enumerate(texts) if phrase in text), None)
        if hit is None:
            missing.append(phrase)
        else:
            found.append({"phrase": phrase, "page": hit})
    result = {"found": found, "missing": missing}
    if args.json:
        rendered = _json_text(result)
    else:
        lines = [f"found: {f['phrase']} (page {f['page']})" for f in found]
        lines += [f"missing: {m}" for m in missing]
        rendered = "\n".join(lines)
    _emit(args, rendered)
    if missing:
        eprint(f"{len(missing)} phrase(s) missing: {', '.join(missing)}")
        return EXIT_FAIL
    return EXIT_OK


# ---------------------------------------------------------------- write commands

def cmd_fields_fill(args) -> int:
    src = Path(args.input_file)
    reader, _ = open_pdf(src, args.password_env)
    out = resolve_target(args, src, reader)
    rows = _field_rows(reader)
    if not rows:
        raise ToolError("This PDF has no form fields.", EXIT_BLOCKED)
    values = _read_json_object(Path(args.values))
    by_name = {r["name"]: r for r in rows}
    unknown = [name for name in values if name not in by_name]
    if unknown:
        raise ToolError(f"Unknown form field(s): {', '.join(unknown)}. No file written.", EXIT_FAIL)

    to_write: dict = {}
    requested: dict = {}
    for name, value in values.items():
        row = by_name[name]
        kind = row["type"]
        if kind == "checkbox":
            if not isinstance(value, bool):
                raise ToolError(f"Value for {name} must be one of: true, false.", EXIT_FAIL)
            on_state = next((s for s in row["_states"] if s != "/Off"), "/Yes")
            to_write[name] = NameObject(on_state if value else "/Off")
        elif kind in ("radio", "dropdown"):
            if not isinstance(value, str) or value not in row["options"]:
                raise ToolError(f"Value for {name} must be one of: {', '.join(row['options'])}.",
                                EXIT_FAIL)
            to_write[name] = NameObject("/" + value) if kind == "radio" else value
        elif kind == "text":
            if not isinstance(value, str):
                raise ToolError(f"Value for {name} must be text.", EXIT_FAIL)
            to_write[name] = value
        else:
            raise ToolError(f"Field {name} is of type {kind} and cannot be filled by this skill.",
                            EXIT_FAIL)
        requested[name] = value

    writer = PdfWriter(clone_from=reader)
    for page in writer.pages:
        if "/Annots" in page:
            writer.update_page_form_field_values(page, to_write, auto_regenerate=False)
    writer.set_need_appearances_writer(True)

    def verify(r) -> None:
        now = {row["name"]: row["value"] for row in _field_rows(r)}
        for name, want in requested.items():
            got = now.get(name)
            if got != want:
                raise VerifyError(f"field {name} reads {got!r}, expected {want!r}")

    _commit(out, lambda fh: writer.write(fh), _pdf_check(verify))
    result = {"out": str(out), "filled": requested, "verified": True}
    return _report_write(args, result, f"Filled {len(requested)} field(s); wrote {out} (verified).")


def cmd_merge(args) -> int:
    inputs = [Path(p) for p in args.inputs]
    if len(inputs) < 2:
        raise ToolError("merge needs at least two PDF inputs.", EXIT_FAIL)
    out = Path(args.out)
    prepare_output(out, args.force, inputs)
    writer = PdfWriter()
    readers, counts = [], []
    for src in inputs:
        reader, _ = open_pdf(src, args.password_env)
        readers.append(reader)
        try:
            counts.append(len(reader.pages))
            for page in reader.pages:
                writer.add_page(page)
        except (PdfReadError, PdfStreamError):
            raise ToolError(corrupt_message(src), EXIT_FAIL) from None
    expected = sum(counts)
    _commit(out, lambda fh: writer.write(fh), _pdf_check(lambda r: _expect_count(r, expected)))
    result = {"out": str(out), "page_count": expected,
              "inputs": [{"path": str(s), "pages": c} for s, c in zip(inputs, counts)],
              "verified": True}
    return _report_write(args, result, f"Merged {len(inputs)} files into {out} ({expected} pages, verified).")


def cmd_split(args) -> int:
    src = Path(args.input_file)
    out_dir = Path(args.out_dir)
    if (args.every is None) == (args.ranges is None):
        raise ToolError("Give exactly one of --every or --ranges.", EXIT_FAIL)
    reader, _ = open_pdf(src, args.password_env)
    total = len(reader.pages)
    if args.every is not None:
        try:
            size = int(args.every)
        except ValueError:
            size = 0
        if size < 1:
            raise ToolError("--every must be a whole number of 1 or more.", EXIT_FAIL)
        groups = [(s, min(s + size - 1, total)) for s in range(1, total + 1, size)]
    else:
        groups = parse_ranges(args.ranges, total)
    if out_dir.exists():
        if not out_dir.is_dir() or any(out_dir.iterdir()):
            raise ToolError(f"Output folder is not empty: {out_dir}", EXIT_FAIL)
    created_folder = not out_dir.exists()
    out_dir.mkdir(parents=True, exist_ok=True)

    staged: list[tuple[Path, Path]] = []
    published: list[Path] = []
    results = []
    try:
        for index, (start, end) in enumerate(groups, start=1):
            writer = PdfWriter()
            for number in range(start, end + 1):
                writer.add_page(reader.pages[number - 1])
            final = out_dir / f"{src.stem}.part{index:02d}.pdf"
            tmp = _stage(out_dir, lambda fh, w=writer: w.write(fh))
            staged.append((tmp, final))
            expected = end - start + 1
            _verify_path(tmp, _pdf_check(lambda r, n=expected: _expect_count(r, n)))
            results.append({"file": final.name, "pages": expected})
        if args.every is not None and sum(r["pages"] for r in results) != total:
            raise ToolError("Output failed verification and was not kept: page totals differ.", EXIT_FAIL)
        for tmp, final in staged:
            _publish(tmp, final)
            published.append(final)
    except BaseException:
        for final in published:
            _discard(final)
        raise
    finally:
        for tmp, _ in staged:
            _discard(tmp)
        if created_folder and not any(out_dir.iterdir()):
            try:
                out_dir.rmdir()
            except OSError:
                pass
    result = {"out_dir": str(out_dir), "parts": results, "verified": True}
    line = f"Wrote {len(results)} part(s) to {out_dir} (verified)."
    return _report_write(args, result, line)


def cmd_rotate(args) -> int:
    if args.angle not in ALLOWED_ANGLES:
        raise ToolError("Angle must be 90, 180 or 270.", EXIT_FAIL)
    angle = int(args.angle)
    src = Path(args.input_file)
    reader, _ = open_pdf(src, args.password_env)
    out = resolve_target(args, src, reader)
    total = len(reader.pages)
    selected = parse_page_list(args.pages, total)
    before =[int(page.get("/Rotate", 0) or 0) for page in reader.pages]
    expected = [((before[i] + angle) % 360) if (i + 1) in selected else before[i]
                for i in range(total)]
    writer = PdfWriter(clone_from=reader)
    for number in selected:
        writer.pages[number - 1][NameObject("/Rotate")] = NumberObject(expected[number - 1])

    def verify(r) -> None:
        _expect_count(r, total)
        for index, page in enumerate(r.pages):
            got = int(page.get("/Rotate", 0) or 0)
            if got != expected[index]:
                raise VerifyError(f"page {index + 1} has rotation {got}, expected {expected[index]}")

    _commit(out, lambda fh: writer.write(fh), _pdf_check(verify))
    result = {"out": str(out), "angle": angle, "rotated_pages": selected, "verified": True}
    return _report_write(args, result, f"Rotated {len(selected)} page(s) by {angle}; wrote {out} (verified).")


def cmd_reorder(args) -> int:
    src = Path(args.input_file)
    reader, _ = open_pdf(src, args.password_env)
    out = resolve_target(args, src, reader)
    total = len(reader.pages)
    message = f"Order must list every page 1..{total} exactly once."
    try:
        order = [int(part) for part in args.order.split(",")]
    except ValueError:
        raise ToolError(message, EXIT_FAIL) from None
    if sorted(order) != list(range(1, total + 1)):
        raise ToolError(message, EXIT_FAIL)
    input_texts = [_norm(page.extract_text() or "") for page in reader.pages]

    writer = PdfWriter(clone_from=reader)
    originals = list(writer.pages)
    for _ in originals:
        writer.remove_page(0)
    for number in order:
        writer.add_page(originals[number - 1])

    def verify(r) -> None:
        _expect_count(r, total)
        for position, number in enumerate(order, start=1):
            got = _norm(r.pages[position - 1].extract_text() or "")
            want = input_texts[number - 1]
            if got and want and got != want:
                raise VerifyError(f"page {position} text does not match input page {number}")

    _commit(out, lambda fh: writer.write(fh), _pdf_check(verify))
    result = {"out": str(out), "order": order, "verified": True}
    return _report_write(args, result, f"Reordered {total} page(s); wrote {out} (verified).")


def _load_reportlab():
    try:
        from reportlab.lib import pagesizes
        from reportlab.pdfbase.pdfmetrics import stringWidth
        from reportlab.pdfgen import canvas
    except ImportError:
        raise ToolError("The text-drawing engine (reportlab) is not installed. "
                        "Run: pip install -r requirements.txt", EXIT_BLOCKED) from None
    return canvas, stringWidth, pagesizes


def _font_has(ch: str) -> bool:
    """True when the base-14 Helvetica font (WinAnsi encoding) can draw the character."""
    if ch == "\n":
        return True
    if ord(ch) < 32 or ord(ch) == 127:
        return False
    try:
        ch.encode("cp1252")
    except UnicodeEncodeError:
        return False
    return True


def _check_font(text: str, skip: str = "\n") -> None:
    line_no = 1
    for ch in text:
        if ch == "\n":
            line_no += 1
            continue
        if ch in skip:
            continue
        if not _font_has(ch):
            raise ToolError(f"Character U+{ord(ch):04X} on line {line_no} is not in the font. "
                            f"No text was drawn.", EXIT_BLOCKED)


def _text_overlay_bytes(text: str, width: float, height: float) -> bytes:
    canvas, string_width, _ = _load_reportlab()
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    lines = text.split("\n")
    diagonal = math.hypot(width, height)
    longest = max(lines, key=lambda s: string_width(s, "Helvetica", 72))
    width72 = string_width(longest, "Helvetica", 72)
    size = min(72.0, 0.8 * diagonal * 72.0 / width72) if width72 > 0 else 72.0
    spacing = size * 1.2
    c.translate(width / 2, height / 2)
    c.rotate(45)
    c.setFillGray(0.5)
    c.setFont("Helvetica", size)
    for i, line in enumerate(lines):
        y = (i - (len(lines) - 1) / 2) * spacing
        c.drawCentredString(0, y - size * 0.35, line)
    c.showPage()
    c.save()
    return buf.getvalue()


def cmd_watermark(args) -> int:
    if (args.stamp is None) == (args.text is None):
        raise ToolError(STAMP_OR_TEXT, EXIT_FAIL)
    src = Path(args.input_file)
    stamp_src = None
    if args.stamp is not None:
        stamp_src = Path(args.stamp)
        if not stamp_src.is_file():
            raise ToolError(STAMP_OR_TEXT, EXIT_FAIL)
    reader, _ = open_pdf(src, args.password_env)
    out = resolve_target(args, src, reader, *([stamp_src] if stamp_src else []))
    total = len(reader.pages)
    selected = parse_page_list(args.pages, total)
    original = [_norm(page.extract_text() or "").split() for page in reader.pages]
    keep_alive = [reader]

    writer = PdfWriter(clone_from=reader)
    if stamp_src is not None:
        stamp_reader, _ = open_pdf(stamp_src, args.password_env)
        keep_alive.append(stamp_reader)
        if not stamp_reader.pages:
            raise ToolError(f"Could not read {stamp_src}: it has no pages.", EXIT_FAIL)
        stamp_page = stamp_reader.pages[0]
        for number in selected:
            writer.pages[number - 1].merge_page(stamp_page)
        method = "stamp"
    else:
        _load_reportlab()
        _check_font(args.text, skip="")
        overlays: dict = {}
        for number in selected:
            page = writer.pages[number - 1]
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            key = (round(width, 2), round(height, 2))
            if key not in overlays:
                overlay_reader = PdfReader(io.BytesIO(_text_overlay_bytes(args.text, width, height)))
                keep_alive.append(overlay_reader)
                overlays[key] = overlay_reader.pages[0]
            page.merge_page(overlays[key])
        method = "text"

    def verify(r) -> None:
        _expect_count(r, total)
        for index, page in enumerate(r.pages):
            got = _norm(page.extract_text() or "").split()
            if not _is_subsequence(original[index], got):
                raise VerifyError(f"page {index + 1} no longer contains its original text")

    _commit(out, lambda fh: writer.write(fh), _pdf_check(verify))
    result = {"out": str(out), "page_count": total, "method": method,
              "watermarked_pages": selected, "verified": True}
    return _report_write(args, result, f"Watermarked {len(selected)} page(s); wrote {out} (verified).")


def cmd_create(args) -> int:
    src, out = Path(args.input_file), Path(args.out)
    if not src.is_file():
        raise ToolError(f"Input file not found: {src}", EXIT_FAIL)
    if out.suffix.lower() != ".pdf":
        raise ToolError(f"Output must end in .pdf: {out}", EXIT_FAIL)
    prepare_output(out, args.force, [src])
    try:
        text = src.read_bytes().decode("utf-8-sig")
    except UnicodeDecodeError:
        raise ToolError(f"Could not read {src}: not UTF-8 text.", EXIT_FAIL) from None
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\t", "    ")
    _check_font(text, skip="\f")
    try:
        font_size = float(args.font_size)
        margin_mm = float(args.margin)
    except ValueError:
        raise ToolError("Font size and margin must be numbers.", EXIT_FAIL) from None
    if font_size <= 0:
        raise ToolError("Font size must be greater than 0.", EXIT_FAIL)
    if margin_mm < 0:
        raise ToolError("Margin must be 0 or more.", EXIT_FAIL)
    _, _, pagesizes = _load_reportlab()
    page_size = pagesizes.A4 if args.page_size == "A4" else pagesizes.letter
    margin = margin_mm * POINTS_PER_MM
    if page_size[0] - 2 * margin <= 0 or page_size[1] - 2 * margin < font_size:
        raise ToolError("Margin is too large for the page.", EXIT_FAIL)

    data = _layout_text_pdf(text, page_size, margin, font_size)
    wanted = [re.sub(r"\s+", "", line) for line in text.split("\n")]
    wanted = [w for w in wanted if w]

    def verify(r) -> None:
        haystack = re.sub(r"\s+", "", "".join(p.extract_text() or "" for p in r.pages))
        pos = 0
        for line in wanted:
            index = haystack.find(line, pos)
            if index < 0:
                raise VerifyError(f"the line '{line[:40]}' is not in the output, in order")
            pos = index + len(line)

    _commit(out, lambda fh: fh.write(data), _pdf_check(verify))
    result = {"out": str(out), "page_count": _page_count_of(data), "verified": True}
    return _report_write(args, result, f"Created {out} ({result['page_count']} page(s), verified).")


def _page_count_of(data: bytes) -> int:
    return len(PdfReader(io.BytesIO(data)).pages)


def _wrap_line(line: str, font: str, size: float, max_width: float, string_width) -> list[str]:
    if string_width(line, font, size) <= max_width:
        return [line]
    pieces: list[str] = []
    current = None
    for word in line.split(" "):
        candidate = word if current is None else current + " " + word
        if string_width(candidate, font, size) <= max_width:
            current = candidate
            continue
        if current is not None:
            pieces.append(current)
            current = None
        piece = ""
        for ch in word:
            if piece and string_width(piece + ch, font, size) > max_width:
                pieces.append(piece)
                piece = ""
            piece += ch
        current = piece
    if current is not None:
        pieces.append(current)
    return pieces or [""]


def _layout_text_pdf(text: str, page_size, margin: float, font_size: float) -> bytes:
    canvas, string_width, _ = _load_reportlab()
    font = "Helvetica"
    leading = font_size * 1.2
    width_pt, height_pt = page_size
    max_width = width_pt - 2 * margin
    top = height_pt - margin - font_size
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=page_size)
    state = {"y": top, "has_content": False, "pending": False, "pages": 0}

    def new_page() -> None:
        c.showPage()
        state["pages"] += 1
        state["y"] = top
        state["has_content"] = False

    def flush_pending() -> None:
        if state["pending"]:
            state["pending"] = False
            new_page()

    def blank_line() -> None:
        flush_pending()
        state["y"] -= leading

    def text_line(line: str) -> None:
        flush_pending()
        for piece in _wrap_line(line, font, font_size, max_width, string_width):
            if state["y"] < margin:
                new_page()
            c.setFont(font, font_size)
            c.drawString(margin, state["y"], piece)
            state["has_content"] = True
            state["y"] -= leading

    def form_feed() -> None:
        if state["has_content"] or state["y"] != top:
            state["pending"] = True

    for line_index, line in enumerate(text.split("\n")):
        for seg_index, segment in enumerate(line.split("\f")):
            if seg_index > 0:
                form_feed()
            if segment.strip() == "":
                blank_line()
            else:
                text_line(segment)
    if state["has_content"] or state["pages"] == 0:
        c.showPage()
    c.save()
    return buf.getvalue()


# ---------------------------------------------------------------- dispatch

def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    parser = build_parser()
    args = parser.parse_args(argv)
    handler = getattr(args, "handler", None)
    if handler is None:
        parser.print_usage(sys.stderr)
        eprint("A COMMAND is required. Run pdf_tool.py --help for the list.")
        return EXIT_FAIL
    try:
        return handler(args)
    except ToolError as exc:
        eprint(exc.message)
        return exc.code
    except (PdfReadError, PdfStreamError):
        hint = getattr(args, "input_file", None) or "the input"
        eprint(corrupt_message(hint))
        return EXIT_FAIL
    except KeyboardInterrupt:
        eprint("Cancelled. No output was kept.")
        return EXIT_FAIL
    except Exception as exc:  # last resort: never print a traceback to the caller
        eprint(f"Internal error ({type(exc).__name__}). No output was kept.")
        return EXIT_FAIL


if __name__ == "__main__":
    sys.exit(main())
