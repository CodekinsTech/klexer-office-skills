#!/usr/bin/env python3
"""xlsx_tool.py - read, write, format, recalculate, check and validate .xlsx workbooks.

One command-line entry point with subcommands:
  sheets, read, write, format, recalc, errors, validate, check-values

Plain-language reference: ../SKILL.md. Libraries, decisions and tests: ../README.md.
"""

from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import math
import os
import posixpath
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import openpyxl
from lxml import etree
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.formula import ArrayFormula

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ERROR_CODES = ("#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!",
               "#VALUE!", "#SPILL!", "#CALC!")
MAX_ROW = 1048576
MAX_COL = 16384
DEFAULT_TIMEOUT = 180.0
OLE_MAGIC = bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])
DATE_PRESET = "yyyy-mm-dd"
PERCENT_PRESET = "0.00%"
SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)

# Functions added after Excel 2007 are stored with the "_xlfn." prefix.
# Written from memory of the Excel function set; NOT checked against Microsoft's
# documentation (see README.md, open items).
XLFN_FUNCTIONS = {
    "AGGREGATE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.INV", "CEILING.MATH",
    "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT",
    "CHISQ.TEST", "CONCAT", "CONFIDENCE.NORM", "CONFIDENCE.T", "COVARIANCE.P",
    "COVARIANCE.S", "DAYS", "ERF.PRECISE", "ERFC.PRECISE", "EXPON.DIST", "F.DIST",
    "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FILTER", "FLOOR.MATH", "FLOOR.PRECISE",
    "FORECAST.LINEAR", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
    "HYPGEOM.DIST", "IFNA", "IFS", "ISFORMULA", "ISOWEEKNUM", "LOGNORM.DIST",
    "LOGNORM.INV", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL", "NEGBINOM.DIST",
    "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
    "NUMBERVALUE", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC",
    "PERCENTRANK.INC", "POISSON.DIST", "QUARTILE.EXC", "QUARTILE.INC", "RANK.AVG",
    "RANK.EQ", "SHEET", "SHEETS", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST",
    "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TEXTJOIN", "UNIQUE",
    "VAR.P", "VAR.S", "WEIBULL.DIST", "WORKDAY.INTL", "XLOOKUP", "XMATCH", "Z.TEST",
    "SEQUENCE", "SORT", "LET",
}
# Functions Excel stores with the "_xlfn._xlws." prefix instead of "_xlfn.".
XLWS_FUNCTIONS = {"FILTER", "SORT"}
CHECK_RULE_KEYS = {"type", "min", "max", "allowed", "pattern", "required", "unique", "max_length"}
CHECK_TYPES = {"integer", "decimal", "date", "boolean", "text"}
PATCH_KEYS = {"sheet", "cell", "value", "formula", "type", "number_format"}
PATCH_TYPES = {"text", "number", "date"}
# Functions that accept an array result as an argument. Every other function takes an array
# only when it has one element (see Evaluator.single_arg).
ARRAY_ARGUMENT_FUNCTIONS = {"SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "TEXTJOIN",
                            "FILTER", "SORT", "UNIQUE"}
# Functions whose result is an array (so a formula that calls one needs array evaluation).
ARRAY_RESULT_FUNCTIONS = {"FILTER", "SORT", "UNIQUE", "SEQUENCE"}
# Error codes that the file format does not have (Excel stores them with dynamic-array metadata,
# which this tool does not write). They are stored as text, so the file opens in Excel.
TEXT_STORED_ERRORS = {"#SPILL!", "#CALC!"}


class ToolError(Exception):
    """An error with the exit code and message the spec requires (message goes to stderr)."""

    def __init__(self, message, exit_code=1):
        super().__init__(message)
        self.exit_code = exit_code


class _Parser(argparse.ArgumentParser):
    """argparse exits 2 on usage errors; the spec requires exit 1, so override it."""

    def error(self, message):
        self.print_usage(sys.stderr)
        sys.stderr.write(f"{self.prog}: error: {message}\n")
        raise SystemExit(1)


# ---------------------------------------------------------------- small helpers

def coord(row, col):
    return f"{get_column_letter(col)}{row}"


def fmt_num(value):
    return int(value) if float(value).is_integer() else value


def num_to_text(x):
    if isinstance(x, bool):
        return "TRUE" if x else "FALSE"
    if isinstance(x, int) or (float(x).is_integer() and abs(x) < 1e15):
        return str(int(x))
    return "%.15g" % x


def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def json_value(v):
    """Convert a cell value to a JSON-safe value (dates as ISO 8601 strings)."""
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, dt.datetime):
        if v.hour or v.minute or v.second or v.microsecond:
            return v.isoformat()
        return v.date().isoformat()
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return str(v)
    return v


def emit(args, payload, text_lines):
    if getattr(args, "json", False):
        sys.stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    else:
        for line in text_lines:
            sys.stdout.write(line + "\n")
    sys.stdout.flush()


def warn(message):
    sys.stderr.write(message + "\n")


def _same_path(a, b):
    return os.path.normcase(os.path.abspath(str(a))) == os.path.normcase(os.path.abspath(str(b)))


# ---------------------------------------------------------------- cell references

_CELL_TEXT_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]+)")


def parse_cell_text(text):
    m = _CELL_TEXT_RE.fullmatch(text.strip())
    if not m:
        raise ToolError(f'"{text}" is not a valid cell or range.')
    col = column_index_from_string(m.group(1).upper())
    row = int(m.group(2))
    if row < 1 or row > MAX_ROW or col > MAX_COL:
        raise ToolError(f'"{text}" is not a valid cell or range.')
    return row, col


def parse_range_text(text):
    """Return (min_row, min_col, max_row, max_col) for 'A1' or 'A1:C10'."""
    parts = text.strip().split(":")
    try:
        if len(parts) == 1:
            r, c = parse_cell_text(parts[0])
            return r, c, r, c
        if len(parts) != 2:
            raise ValueError
        r1, c1 = parse_cell_text(parts[0])
        r2, c2 = parse_cell_text(parts[1])
    except (ToolError, ValueError):
        raise ToolError(f'"{text}" is not a valid cell or range.') from None
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def pick_sheet(wb, name):
    if not wb.worksheets:
        raise ToolError("The workbook has no sheets.")
    if name is None:
        for ws in wb.worksheets:
            if ws.sheet_state == "visible":
                return ws
        return wb.worksheets[0]
    if name in wb.sheetnames:
        return wb[name]
    raise ToolError(f'No sheet named "{name}". Sheets are: {", ".join(wb.sheetnames)}.')


def used_range(ws):
    return ws.min_row, ws.min_column, ws.max_row, ws.max_column


# ---------------------------------------------------------------- input checks

def check_format(path):
    """Refuse legacy, macro-enabled, CSV/TSV and encrypted inputs (exit 2)."""
    p = Path(path)
    suffix = p.suffix.lower()
    if suffix == ".xls":
        raise ToolError("Legacy .xls is not supported. Save as .xlsx and try again.", 2)
    if suffix == ".xlsm":
        raise ToolError("Macro-enabled files are not supported. Save as .xlsx (macros would be dropped).", 2)
    if suffix in (".csv", ".tsv"):
        raise ToolError("CSV and TSV are not workbooks. Use a workbook, or ask for a CSV tool.", 2)
    with open(p, "rb") as fh:
        head = fh.read(8)
    if head == OLE_MAGIC:
        if suffix == ".xlsx":
            raise ToolError("This workbook is password-protected. The skill does not open protected files.", 2)
        raise ToolError("Legacy .xls is not supported. Save as .xlsx and try again.", 2)
    if suffix != ".xlsx":
        raise ToolError(f"Not an .xlsx file: {path}. Save as .xlsx and try again.")


def check_input(path):
    if not Path(path).is_file():
        raise ToolError(f"Input file not found: {path}")
    check_format(path)


def open_workbook(path, data_only=False):
    check_input(path)
    if not zipfile.is_zipfile(path):
        raise ToolError(f"File is not a readable .xlsx (corrupt or truncated): {path}")
    try:
        return openpyxl.load_workbook(str(path), data_only=data_only)
    except PermissionError:
        raise
    except Exception:
        raise ToolError(f"File is not a readable .xlsx (corrupt or truncated): {path}") from None


def pillow_available():
    try:
        import PIL  # noqa: F401
        return True
    except ImportError:
        return False


_SHAPE_TAGS = {"graphicFrame", "sp", "grpSp", "cxnSp", "contentPart"}


def drawing_status(path):
    """Return (has_charts_or_shapes, has_pictures) by reading the drawing parts."""
    has_drawing = False
    has_picture = False
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if name.startswith("xl/charts/") and name.endswith(".xml"):
                has_drawing = True
            elif re.fullmatch(r"xl/drawings/[^/]+\.xml", name):
                try:
                    root = etree.fromstring(zf.read(name), SAFE_PARSER)
                except etree.XMLSyntaxError:
                    has_drawing = True
                    continue
                for el in root.iter():
                    if not isinstance(el.tag, str):
                        continue
                    local = etree.QName(el).localname
                    if local in _SHAPE_TAGS:
                        has_drawing = True
                    elif local == "pic":
                        has_picture = True
    return has_drawing, has_picture


def drawing_guard(path, allow):
    """Apply the drawing guard. Returns True when drawings may be rewritten (allowed)."""
    has_drawing, has_picture = drawing_status(path)
    if has_drawing:
        if not allow:
            raise ToolError("This workbook has charts or drawing shapes that may be altered on save. "
                            "Pass --allow-drawing-rewrite to continue.", 2)
        return True
    if has_picture and not pillow_available():
        if not allow:
            raise ToolError("This workbook has pictures, and Pillow is not installed, so they would be "
                            "dropped on save. Install Pillow or pass --allow-drawing-rewrite to continue.", 2)
        return True
    return False


def plan_target(src, out, in_place, tag, force, require_xlsx=True):
    """Work out the output path, or return None when no output is requested."""
    if in_place and out:
        raise ToolError("Give either --out or --in-place, not both.")
    if in_place:
        return Path(src)
    if out:
        target = Path(out)
    elif tag:
        src_path = Path(src)
        target = src_path.with_name(f"{src_path.stem}.{tag}.xlsx")
    else:
        return None
    if require_xlsx and target.suffix.lower() != ".xlsx":
        raise ToolError(f"Output must be an .xlsx file: {target}")
    if _same_path(target, src):
        raise ToolError("Refusing to overwrite the input. Use --in-place to change it.")
    if not target.parent.is_dir():
        raise ToolError(f"Output folder does not exist: {target.parent}")
    if target.exists() and not force:
        raise ToolError(f"Output already exists: {target}. Choose another --out or pass --force.")
    return target


def commit_workbook(save_fn, target, verify_fn):
    """Save to a temp file in the output folder, verify it, then move it into place."""
    fd, tmp_name = tempfile.mkstemp(prefix=".xlsx-tool-", suffix=".xlsx", dir=str(target.parent))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        save_fn(tmp)
        try:
            verify_fn(tmp)
        except ToolError as exc:
            raise ToolError(f"Output failed verification and was not kept: {exc}") from None
        os.replace(tmp, target)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


def write_text_output(target, text):
    fd, tmp_name = tempfile.mkstemp(prefix=".xlsx-tool-", suffix=".tmp", dir=str(target.parent))
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        json.loads(tmp.read_text(encoding="utf-8"))
        os.replace(tmp, target)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


def values_equal(expected, got):
    if expected is None:
        return got is None or got == ""
    if isinstance(expected, dt.datetime):
        return isinstance(got, dt.datetime) and expected == got
    if isinstance(expected, bool):
        return isinstance(got, bool) and expected == got
    if is_number(expected):
        return is_number(got) and math.isclose(float(expected), float(got), rel_tol=1e-9, abs_tol=1e-12)
    return expected == got


# ---------------------------------------------------------------- xlfn prefix

_XLFN_RE = re.compile(r"\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'|(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)(?=\()")


def add_xlfn_prefix(formula):
    """Add _xlfn. to post-2007 function names, and _xlpm. to LET names.

    Returns (new_formula, [function names])."""
    found = []

    def repl(m):
        name = m.group(1)
        if name is None:
            return m.group(0)
        upper = name.upper()
        if upper in XLFN_FUNCTIONS:
            found.append(upper)
            prefix = "_xlfn._xlws." if upper in XLWS_FUNCTIONS else "_xlfn."
            return prefix + name
        return name

    new_formula = _XLFN_RE.sub(repl, formula)
    if "LET" in found:
        new_formula = _prefix_let_names(new_formula)
    return new_formula, found


# Tokens for the LET rewrite: strings, quoted sheet names, words, whitespace, other characters.
_SCAN_RE = re.compile(r"""(?P<ws>\s+)|(?P<str>"(?:[^"]|"")*")|(?P<sq>'(?:[^']|'')*')"""
                      r"""|(?P<word>[A-Za-z_\\][A-Za-z0-9_.]*)|(?P<other>.)""", re.S)


def _strip_xl(word):
    """Remove a leading _xlfn., _xlws. or _xlpm. marker from a word."""
    lower = word.lower()
    for marker in ("_xlfn.", "_xlws.", "_xlpm."):
        if lower.startswith(marker):
            return word[len(marker):]
    return word


def _sig_next(toks, i):
    """Index of the next token after i that is not whitespace, or None."""
    j = i + 1
    while j < len(toks) and toks[j][0] == "ws":
        j += 1
    return j if j < len(toks) else None


def _sig_prev(toks, i):
    j = i - 1
    while j >= 0 and toks[j][0] == "ws":
        j -= 1
    return j if j >= 0 else None


def _prefix_let_names(formula):
    """Give each LET parameter name, and each use of one inside the LET call, the _xlpm. prefix."""
    toks = [[m.lastgroup, m.group(0)] for m in _SCAN_RE.finditer(formula)]
    i = 0
    while i < len(toks):
        nxt = _sig_next(toks, i)
        if toks[i][0] == "word" and _strip_xl(toks[i][1]).upper() == "LET" \
                and nxt is not None and toks[nxt][1] == "(":
            # find the matching close parenthesis and the top-level argument boundaries
            depth = 0
            args = [[]]
            close = None
            for k in range(nxt, len(toks)):
                kind, text = toks[k]
                if kind == "other" and text == "(":
                    depth += 1
                    if depth == 1:
                        continue
                elif kind == "other" and text == ")":
                    depth -= 1
                    if depth == 0:
                        close = k
                        break
                elif kind == "other" and text == "," and depth == 1:
                    args.append([])
                    continue
                if depth >= 1 and kind != "ws":
                    args[-1].append(k)
            if close is None:
                break
            names = set()
            for n, arg in enumerate(args[:-1]):
                if n % 2 == 0 and len(arg) == 1 and toks[arg[0]][0] == "word":
                    names.add(_strip_xl(toks[arg[0]][1]).upper())
            for k in range(nxt + 1, close):
                if toks[k][0] != "word":
                    continue
                base = _strip_xl(toks[k][1])
                if base.upper() not in names or toks[k][1].lower().startswith("_xlpm."):
                    continue
                after = _sig_next(toks, k)
                before = _sig_prev(toks, k)
                if after is not None and toks[after][1] == "(":
                    continue
                if before is not None and toks[before][1] == "!":
                    continue
                toks[k][1] = "_xlpm." + base
        i += 1
    return "".join(text for _, text in toks)


# ---------------------------------------------------------------- patch file

def load_patch(path):
    p = Path(path)
    if not p.is_file():
        raise ToolError(f"Patch file not found: {path}")
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        raise ToolError(f"Patch file is malformed: {exc}.") from None
    if not isinstance(data, list):
        raise ToolError("Patch file is malformed: it must contain a JSON list of entries.")
    for n, entry in enumerate(data, 1):

        def bad(reason, n=n):
            raise ToolError(f"Patch entry {n} is invalid: {reason}.")

        if not isinstance(entry, dict):
            bad("it is not an object")
        for key in entry:
            if key not in PATCH_KEYS:
                bad(f'unknown key "{key}"')
        if not isinstance(entry.get("sheet"), str) or not entry.get("sheet"):
            bad('"sheet" must be a sheet name')
        if not isinstance(entry.get("cell"), str):
            bad('"cell" must be a cell reference such as "B2"')
        has_value = "value" in entry
        has_formula = "formula" in entry
        if has_value == has_formula:
            bad('give exactly one of "value" or "formula"')
        entry_type = entry.get("type")
        if entry_type is not None and entry_type not in PATCH_TYPES:
            bad('"type" must be "text", "number" or "date"')
        nf = entry.get("number_format")
        if nf is not None and (not isinstance(nf, str) or not nf):
            bad('"number_format" must be a non-empty text code')
        if has_formula:
            if not isinstance(entry["formula"], str):
                bad('"formula" must be text')
            if entry_type is not None:
                bad('"type" cannot be used with "formula"')
        else:
            value = entry["value"]
            if isinstance(value, (dict, list)):
                bad("value must be a number, text, boolean or null")
            if isinstance(value, str) and value.startswith("="):
                bad('text that starts with "=" must be given as "formula"')
            if entry_type == "number" and isinstance(value, str):
                if not re.fullmatch(r"\s*[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?\s*", value):
                    bad(f'"{value}" is not a number')
            if entry_type == "number" and isinstance(value, bool):
                bad("a boolean cannot be written as a number")
    return data


def convert_patch_value(entry, n, ref):
    entry_type = entry.get("type")
    value = entry.get("value")
    if value is None:
        return None
    if entry_type == "date":
        if not (isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)):
            raise ToolError(f'"{value}" is not a valid YYYY-MM-DD date (cell {ref}).')
        try:
            return dt.datetime.strptime(value, "%Y-%m-%d")
        except ValueError:
            raise ToolError(f'"{value}" is not a valid YYYY-MM-DD date (cell {ref}).') from None
    if entry_type == "number" and isinstance(value, str):
        text = value.strip()
        return float(text) if re.search(r"[.eE]", text) else int(text)
    if entry_type == "text":
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        if is_number(value):
            return num_to_text(value)
    return value


# ---------------------------------------------------------------- formula evaluator

class Unsupported(Exception):
    """The built-in evaluator does not implement something the formula uses."""

    def __init__(self, reason):
        super().__init__(reason)
        self.reason = reason
        self.where = None


class Circular(Exception):
    def __init__(self, key):
        super().__init__(key)
        self.key = key


class Timeout(Exception):
    pass


class XLError:
    """A spreadsheet error value such as #DIV/0!."""
    __slots__ = ("code",)

    def __init__(self, code):
        self.code = code

    def __eq__(self, other):
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self):
        return hash(self.code)

    def __repr__(self):
        return f"XLError({self.code!r})"


DIV0 = XLError("#DIV/0!")
VALUE = XLError("#VALUE!")
NA = XLError("#N/A")
REF = XLError("#REF!")
NUM = XLError("#NUM!")
NAME = XLError("#NAME?")
SPILL = XLError("#SPILL!")
CALC = XLError("#CALC!")


class Arr:
    """A two-dimensional array result from FILTER, SORT, UNIQUE, SEQUENCE or LET (at least 1x1)."""
    __slots__ = ("rows",)

    def __init__(self, rows):
        self.rows = rows

    @property
    def height(self):
        return len(self.rows)

    @property
    def width(self):
        return len(self.rows[0])

    def transposed(self):
        return Arr([list(line) for line in zip(*self.rows)])


class Outcome:
    """The result of one formula cell. block is the spilled Arr when the formula returned several
    values and the spill was placed; value is then the top-left value (or #SPILL!). array is True
    when the formula needed array evaluation (an operator on a range or array, or FILTER, SORT,
    UNIQUE, SEQUENCE or an XLOOKUP that returns a range). Such a single-cell formula is saved as a
    one-cell array formula, so Excel calculates it the same way when the file is opened."""
    __slots__ = ("value", "block", "array")

    def __init__(self, value, block=None, array=False):
        self.value = value
        self.block = block
        self.array = array


_LET_NAME_RE = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.]*")


def valid_let_name(name):
    """A LET name must start with a letter, underscore or backslash, and must not look like a cell
    reference. Excel (checked in Microsoft 365) accepts R, C and RC as names, so they are allowed.
    It rejects R1, C1, RC1, R100 and similar, so those are refused here too. Names with digits
    in R1C1 form (R1C1, R2C3) and A1-style names are refused: Excel gives inconsistent results
    for them inside a LET, so no value is guessed."""
    if not _LET_NAME_RE.fullmatch(name):
        return False
    if re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[0-9]+", name):
        return False
    if re.fullmatch(r"(?i)(R|C|RC)\.?[0-9]+\.?", name):
        return False
    if re.fullmatch(r"(?i)(?=.*[0-9])R[0-9]*C[0-9]*", name):
        return False
    return True


class Ref:
    """A rectangular reference on one sheet (1-based, inclusive)."""
    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet = sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2

    @property
    def height(self):
        return self.r2 - self.r1 + 1

    @property
    def width(self):
        return self.c2 - self.c1 + 1


_TOKEN_RE = re.compile(r"""
    (?P<ws>\s+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<err>\#(?:DIV/0!|N/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!))
  | (?P<sheet>(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)
  | (?P<func>(?:_xlfn\.)?[A-Za-z_][A-Za-z0-9_.]*(?=\())
  | (?P<ref>\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?)(?![A-Za-z0-9_.])
  | (?P<colref>\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![A-Za-z0-9_.])
  | (?P<num>(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)
  | (?P<bool>(?:TRUE|FALSE)(?![A-Za-z0-9_.(]))
  | (?P<name>[A-Za-z_][A-Za-z0-9_.]*)
  | (?P<op><>|<=|>=|[-+*/^&=<>%(),])
""", re.VERBOSE)

_CELL_PART_RE = re.compile(r"([A-Za-z]{1,3})([0-9]+)")
_NUM_TEXT_RE = re.compile(r"[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")
_NAME_TARGET_RE = re.compile(r"(?:'((?:[^']|'')+)'|([^'!\[\]]+))!(.+)")
_COL_ONLY_RE = re.compile(r"[A-Za-z]{1,3}")


def _split_cell(text):
    m = _CELL_PART_RE.fullmatch(text)
    if not m:
        raise Unsupported(f"bad cell reference {text}")
    row = int(m.group(2))
    col = column_index_from_string(m.group(1).upper())
    if row < 1 or row > MAX_ROW or col > MAX_COL:
        raise Unsupported(f"cell reference {text} is outside the sheet")
    return row, col


def _parse_ref_text(text):
    parts = text.replace("$", "").split(":")
    r1, c1 = _split_cell(parts[0])
    if len(parts) == 2:
        r2, c2 = _split_cell(parts[1])
    else:
        r2, c2 = r1, c1
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def tokenize(text):
    tokens = []
    pos = 0
    while pos < len(text):
        m = _TOKEN_RE.match(text, pos)
        if not m:
            raise Unsupported(f"unexpected text at character {pos + 1}")
        kind = m.lastgroup
        tokens.append((kind, m.group(kind)))
        pos = m.end()
    return [t for t in tokens if t[0] != "ws"]


class _FormulaParser:
    """Recursive-descent parser with Excel's operator precedence."""

    def __init__(self, tokens):
        self.toks = tokens
        self.pos = 0

    def peek(self):
        return self.toks[self.pos] if self.pos < len(self.toks) else (None, None)

    def take(self):
        tok = self.peek()
        self.pos += 1
        return tok

    def is_op(self, *ops):
        kind, val = self.peek()
        return kind == "op" and val in ops

    def parse(self):
        node = self.comparison()
        if self.pos != len(self.toks):
            raise Unsupported("unexpected text after the formula")
        return node

    def comparison(self):
        node = self.concat()
        while self.is_op("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            node = ("bin", op, node, self.concat())
        return node

    def concat(self):
        node = self.additive()
        while self.is_op("&"):
            self.take()
            node = ("bin", "&", node, self.additive())
        return node

    def additive(self):
        node = self.term()
        while self.is_op("+", "-"):
            op = self.take()[1]
            node = ("bin", op, node, self.term())
        return node

    def term(self):
        node = self.power()
        while self.is_op("*", "/"):
            op = self.take()[1]
            node = ("bin", op, node, self.power())
        return node

    def power(self):
        node = self.unary()
        while self.is_op("^"):
            self.take()
            node = ("bin", "^", node, self.unary())
        return node

    def unary(self):
        if self.is_op("-"):
            self.take()
            return ("neg", self.unary())
        if self.is_op("+"):
            self.take()
            return self.unary()
        return self.postfix()

    def postfix(self):
        node = self.primary()
        while True:
            if self.is_op("%"):
                self.take()
                node = ("pct", node)
            elif self.is_op("("):
                # A value followed by ( is a call of that value, as in LAMBDA(x,x*2)(5).
                self.take()
                args = []
                if not self.is_op(")"):
                    args.append(self.comparison())
                    while self.is_op(","):
                        self.take()
                        args.append(self.comparison())
                k2, v2 = self.take()
                if not (k2 == "op" and v2 == ")"):
                    raise Unsupported("missing closing parenthesis")
                node = ("invoke", node, args)
            else:
                return node

    def primary(self):
        kind, val = self.take()
        if kind is None:
            raise Unsupported("the formula ends too early")
        if kind == "num":
            return ("lit", float(val) if re.search(r"[.eE]", val) else int(val))
        if kind == "str":
            return ("lit", val[1:-1].replace('""', '"'))
        if kind == "bool":
            return ("lit", val.upper() == "TRUE")
        if kind == "err":
            return ("lit", XLError(val))
        if kind == "ref":
            return ("ref", None) + _parse_ref_text(val)
        if kind == "colref":
            c1, c2 = (column_index_from_string(p.replace("$", "").upper()) for p in val.split(":"))
            return ("colref", None, min(c1, c2), max(c1, c2))
        if kind == "sheet":
            name = val[:-1]
            if name.startswith("'"):
                name = name[1:-1].replace("''", "'")
            k2, v2 = self.take()
            if k2 == "ref":
                return ("ref", name) + _parse_ref_text(v2)
            if k2 == "colref":
                c1, c2 = (column_index_from_string(p.replace("$", "").upper()) for p in v2.split(":"))
                return ("colref", name, min(c1, c2), max(c1, c2))
            raise Unsupported("a sheet name must be followed by a cell or range")
        if kind == "func":
            name = val.upper()
            while name.startswith(("_XLFN.", "_XLWS.")):
                name = name.split(".", 1)[1]
            if not re.fullmatch(r"[A-Z][A-Z0-9]*", name):
                raise Unsupported(f"the function {name} is not supported")
            k2, v2 = self.take()
            if not (k2 == "op" and v2 == "("):
                raise Unsupported("a function name must be followed by (")
            args = []
            if self.is_op(")"):
                self.take()
                return ("func", name, args)
            while True:
                if self.is_op(",", ")"):
                    args.append(("lit", None))
                else:
                    args.append(self.comparison())
                k3, v3 = self.take()
                if k3 == "op" and v3 == ",":
                    continue
                if k3 == "op" and v3 == ")":
                    break
                raise Unsupported("bad function arguments")
            return ("func", name, args)
        if kind == "op" and val == "(":
            node = self.comparison()
            k2, v2 = self.take()
            if not (k2 == "op" and v2 == ")"):
                raise Unsupported("missing closing parenthesis")
            return node
        if kind == "name":
            if val.lower().startswith("_xlpm."):
                val = val[6:]
            return ("name", val)
        raise Unsupported(f"syntax error near {val!r}")


def parse_formula(text):
    tokens = tokenize(text)
    if not tokens:
        raise Unsupported("the formula is empty")
    return _FormulaParser(tokens).parse()


def to_number(v):
    if v is None:
        return 0
    if isinstance(v, XLError):
        return v
    if isinstance(v, bool):
        return 1 if v else 0
    if is_number(v):
        return v
    if isinstance(v, str) and _NUM_TEXT_RE.fullmatch(v.strip()):
        return float(v.strip())
    return VALUE


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if is_number(v):
        return num_to_text(v)
    return str(v)


def to_bool(v):
    if v is None:
        return False
    if isinstance(v, XLError):
        return v
    if isinstance(v, bool):
        return v
    if is_number(v):
        return v != 0
    if isinstance(v, str):
        upper = v.strip().upper()
        if upper == "TRUE":
            return True
        if upper == "FALSE":
            return False
    return VALUE


def _rank(v):
    if isinstance(v, bool):
        return 2
    if is_number(v):
        return 0
    if isinstance(v, str):
        return 1
    return None


def _order_key(v):
    return v.casefold() if isinstance(v, str) else v


def _lookup_key(v):
    return 0 if v is None else v


def _loose_eq(a, b):
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if is_number(a) and is_number(b):
        return a == b
    if isinstance(a, str) and isinstance(b, str):
        return a.casefold() == b.casefold()
    return False


def _blank_like(other):
    if isinstance(other, str):
        return ""
    if isinstance(other, bool):
        return False
    return 0


def compare_values(op, a, b):
    if a is None:
        a = _blank_like(b)
    if b is None:
        b = _blank_like(a)
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        c = (ra > rb) - (ra < rb)
    else:
        ka, kb = _order_key(a), _order_key(b)
        c = (ka > kb) - (ka < kb)
    return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]


_WILDCARD_CHARS = set("*?")


def wildcard_regex(pattern):
    out = []
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    return re.compile("".join(out), re.S | re.I)


def make_predicate(crit):
    """Turn a COUNTIF-style criterion into a test on one cell value."""
    if crit is None:
        crit = ""
    if isinstance(crit, bool):
        return lambda v: isinstance(v, bool) and v == crit
    if is_number(crit):
        return lambda v: is_number(v) and v == crit
    text = str(crit)
    m = re.fullmatch(r"(<=|>=|<>|<|>|=)?(.*)", text, re.S)
    op, operand = m.group(1), m.group(2)
    if operand == "" and op in (None, "="):
        return lambda v: v is None or v == ""
    if operand == "" and op == "<>":
        return lambda v: not (v is None or v == "")
    if _NUM_TEXT_RE.fullmatch(operand):
        x = float(operand)
        if op in (None, "="):
            return lambda v: is_number(v) and v == x
        if op == "<>":
            return lambda v: not (is_number(v) and v == x)
        ops = {"<": lambda a, b: a < b, ">": lambda a, b: a > b,
               "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b}[op]
        return lambda v: is_number(v) and ops(v, x)
    if operand.upper() in ("TRUE", "FALSE") and op in (None, "="):
        want = operand.upper() == "TRUE"
        return lambda v: isinstance(v, bool) and v == want
    pat = wildcard_regex(operand)
    if op in (None, "="):
        return lambda v: isinstance(v, str) and pat.fullmatch(v) is not None
    if op == "<>":
        return lambda v: not (isinstance(v, str) and pat.fullmatch(v) is not None)
    ops = {"<": lambda a, b: a < b, ">": lambda a, b: a > b,
           "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b}[op]
    folded = operand.casefold()
    return lambda v: isinstance(v, str) and ops(v.casefold(), folded)


def _has_wildcard(text):
    return any(ch in _WILDCARD_CHARS for ch in text) or "~" in text


def _ordered_pairs(pairs, direction):
    """Return (index, value) pairs in ascending order for a binary search.

    direction 1: the values must already ascend. direction -1: they must descend, and the
    pairs are reversed. Unsorted values are refused, because Excel's answer for them is not
    documented."""
    keys = [_order_key(v) for _, v in pairs]
    if direction == 1:
        in_order = all(a <= b for a, b in zip(keys, keys[1:]))
    else:
        in_order = all(a >= b for a, b in zip(keys, keys[1:]))
    if not in_order:
        raise Unsupported("the lookup values are not sorted in the order this search needs "
                          "(binary search on unsorted data is not implemented)")
    return pairs if direction == 1 else pairs[::-1]


def _count_true(n, pred):
    """Count the leading True results of pred(0..n-1). pred must be True first, then False."""
    lo, hi = 0, n
    while lo < hi:
        mid = (lo + hi) // 2
        if pred(mid):
            lo = mid + 1
        else:
            hi = mid
    return lo


def _bsearch_pick(keys, key, mode):
    """Binary search over keys sorted ascending. Return the chosen position, or None.

    mode: "exact" (XLOOKUP match_mode 0), "largest_le" (MATCH and VLOOKUP approximate),
    "smallest_ge" (MATCH -1 on descending data, searched after reversal), "next_smaller"
    (XLOOKUP match_mode -1) and "next_larger" (XLOOKUP match_mode 1).
    A chosen value that occurs more than once is refused, because which copy Excel returns
    is not documented."""
    n = len(keys)
    lb = _count_true(n, lambda i: keys[i] < key)    # number of keys below key
    ub = _count_true(n, lambda i: keys[i] <= key)   # number of keys at or below key
    dup_msg = ("the chosen value occurs more than once in the lookup range; which copy Excel "
               "returns is not documented, so this case is not implemented")
    if mode == "exact":
        if ub == lb:
            return None
        if ub - lb > 1:
            raise Unsupported(dup_msg)
        return lb
    if mode in ("next_smaller", "next_larger") and ub > lb:
        # match_mode -1 and 1 take an exact match first (Excel: "exact match or next smaller/larger")
        if ub - lb > 1:
            raise Unsupported(dup_msg)
        return lb
    if mode == "largest_le":
        if ub == 0:
            return None
        pos = ub - 1
        dup = pos > 0 and keys[pos - 1] == keys[pos]
    elif mode == "smallest_ge":
        if lb == n:
            return None
        pos = lb
        dup = pos + 1 < n and keys[pos + 1] == keys[pos]
    elif mode == "next_smaller":
        if lb == 0:
            return None
        pos = lb - 1
        dup = pos > 0 and keys[pos - 1] == keys[pos]
    elif mode == "next_larger":
        if ub == n:
            return None
        pos = ub
        dup = pos + 1 < n and keys[pos + 1] == keys[pos]
    else:
        raise ValueError(mode)
    if dup:
        raise Unsupported(dup_msg)
    return pos


class Evaluator:
    """Evaluates the formula functions listed in the spec, one cell at a time, with memo.

    names maps (sheet or None, NAME in upper case) to the stored text of a defined name
    (None = workbook scope, a sheet title = sheet scope)."""

    def __init__(self, model, deadline, names=None, candidates=None, array_anchors=None):
        self.model = model
        self.deadline = deadline
        self.outcomes = {}
        self.active = set()
        self.ast_cache = {}
        self.names = names or {}
        self.name_cache = {}
        self.scopes = []          # LET frames, innermost last; reset for every formula cell
        self.array_used = False   # set when the formula being calculated needs array evaluation
        # sheet -> sorted (row, col) of formula cells that may return an array
        self.candidates = candidates or {}
        # formula cells of multi-cell array formulas that are not array functions (refused if scalar)
        self.array_anchors = array_anchors or set()
        self.claims = {}          # (sheet, row, col) of a spilled cell -> (anchor key, i, j)
        self.pending = {}         # anchor key -> empty cells read while the anchor was calculating

    # -- cells and references
    def max_row_of(self, sheet):
        sm = self.model.get(sheet)
        if sm is None:
            raise Unsupported(f'reference to an unknown sheet "{sheet}"')
        return sm.max_row

    def cell_value(self, sheet, row, col):
        sm = self.model.get(sheet)
        if sm is None:
            raise Unsupported(f'reference to an unknown sheet "{sheet}"')
        entry = sm.cells.get((row, col))
        if entry is None:
            return self.owned_value(sheet, row, col)
        if entry[0] == "v":
            return entry[1]
        return self.evaluate_formula((sheet, row, col)).value

    def owned_value(self, sheet, row, col):
        """Value of a cell with no entry of its own. It may lie inside the spilled block of a
        formula that is up and to the left of it, so those formulas are calculated first."""
        for crow, ccol in self.candidates.get(sheet, ()):
            if crow > row or ccol > col:
                continue
            key = (sheet, crow, ccol)
            if key in self.active:
                # the formula is still calculating; if its block turns out to cover this cell
                # it depends on itself, which evaluate_formula reports as a circular reference
                self.pending.setdefault(key, set()).add((sheet, row, col))
                continue
            self.evaluate_formula(key)
        hit = self.claims.get((sheet, row, col))
        if hit is None:
            return None
        anchor, i, j = hit
        return self.outcomes[anchor].block.rows[i][j]

    def evaluate_formula(self, key):
        """Calculate one formula cell once and remember the outcome."""
        done = self.outcomes.get(key)
        if done is not None:
            return done
        if key in self.active:
            raise Circular(key)
        if time.monotonic() > self.deadline:
            raise Timeout()
        sheet, row, col = key
        self.active.add(key)
        saved_scopes, self.scopes = self.scopes, []
        saved_array, self.array_used = self.array_used, False
        try:
            ast = self.ast_cache.get(key)
            if ast is None:
                ast = parse_formula(self.model[sheet].cells[(row, col)][1])
                self.ast_cache[key] = ast
            result = self.ev(ast, sheet)
            uses_array = self.array_used
            if isinstance(result, Ref):
                result = self.single(result)
            if isinstance(result, Arr) and result.height == 1 and result.width == 1:
                result = result.rows[0][0]
            if key in self.array_anchors and not isinstance(result, Arr):
                raise Unsupported("a multi-cell array formula that returns one value is not calculated")
            if isinstance(result, Arr):
                outcome = self.spill_outcome(key, result, self.pending.pop(key, set()))
            else:
                if result is None:
                    result = 0
                if isinstance(result, float) and (math.isnan(result) or math.isinf(result)):
                    result = NUM
                outcome = Outcome(result, array=uses_array)
        except Unsupported as exc:
            if exc.where is None:
                exc.where = key
            raise
        finally:
            self.active.discard(key)
            self.scopes = saved_scopes
            self.array_used = saved_array
        self.outcomes[key] = outcome
        if outcome.block is not None:
            sheet, row, col = key
            for i in range(outcome.block.height):
                for j in range(outcome.block.width):
                    if i or j:
                        self.claims[(sheet, row + i, col + j)] = (key, i, j)
        return outcome

    def spill_outcome(self, key, arr, pending):
        """Place an array result. The result is refused with #SPILL! when it would run past the
        sheet, cover a non-empty cell, or overlap a range that an earlier formula already claimed."""
        sheet, row, col = key
        block = Arr([[0 if v is None else v for v in line] for line in arr.rows])
        height, width = block.height, block.width
        if row + height - 1 > MAX_ROW or col + width - 1 > MAX_COL:
            return Outcome(SPILL, array=True)
        for psheet, pr, pc in pending:
            if psheet == sheet and row <= pr < row + height and col <= pc < col + width:
                raise Circular(key)
        cells = self.model[sheet].cells
        for i in range(height):
            for j in range(width):
                if i == 0 and j == 0:
                    continue
                if (row + i, col + j) in cells or (sheet, row + i, col + j) in self.claims:
                    return Outcome(SPILL, array=True)
        return Outcome(block.rows[0][0], block)

    def iter_ref(self, ref):
        for r in range(ref.r1, ref.r2 + 1):
            for c in range(ref.c1, ref.c2 + 1):
                yield self.cell_value(ref.sheet, r, c)

    def single(self, ref):
        if ref.height == 1 and ref.width == 1:
            return self.cell_value(ref.sheet, ref.r1, ref.c1)
        raise Unsupported("a range is used where a single value is needed (array or implicit intersection)")

    def scalar(self, v):
        if isinstance(v, Ref):
            return self.single(v)
        if isinstance(v, Arr):
            if v.height == 1 and v.width == 1:
                return v.rows[0][0]
            raise Unsupported("an array is used where a single value is needed")
        return v

    # -- defined names
    def resolve_name(self, name, sheet):
        upper = name.upper()
        scope = (sheet, upper) if (sheet, upper) in self.names else (None, upper)
        if scope not in self.names:
            return NAME  # not defined anywhere in the workbook: Excel gives #NAME?
        if scope not in self.name_cache:
            self.name_cache[scope] = self.name_target(name, self.names[scope])
        return self.name_cache[scope]

    def name_target(self, name, text):
        """Turn a defined name's stored text (for example "Data!$A$1:$A$3") into a Ref."""
        text = (text or "").strip()
        if text == "#REF!":
            return REF
        if text.startswith("["):
            raise Unsupported(f'the name "{name}" refers to another workbook')
        m = _NAME_TARGET_RE.fullmatch(text)
        if m is None:
            raise Unsupported(f'the name "{name}" does not refer to a single cell or range')
        sheet = m.group(1).replace("''", "'") if m.group(1) is not None else m.group(2)
        body = m.group(3)
        if sheet not in self.model:
            return REF
        parts = body.replace("$", "").split(":")
        if len(parts) == 2 and all(_COL_ONLY_RE.fullmatch(p) for p in parts):
            c1, c2 = (column_index_from_string(p.upper()) for p in parts)
            return Ref(sheet, 1, min(c1, c2), self.max_row_of(sheet), max(c1, c2))
        try:
            r1, c1, r2, c2 = _parse_ref_text(body)
        except Unsupported:
            raise Unsupported(f'the name "{name}" does not refer to a single cell or range') from None
        return Ref(sheet, r1, c1, r2, c2)

    # -- expression evaluation
    def ev(self, node, sheet):
        kind = node[0]
        if kind == "lit":
            return node[1]
        if kind == "ref":
            s = node[1] or sheet
            if s not in self.model:
                raise Unsupported(f'reference to an unknown sheet "{s}"')
            return Ref(s, node[2], node[3], node[4], node[5])
        if kind == "colref":
            s = node[1] or sheet
            return Ref(s, 1, node[2], self.max_row_of(s), node[3])
        if kind == "neg":
            x = to_number(self.scalar(self.ev(node[1], sheet)))
            return x if isinstance(x, XLError) else -x
        if kind == "pct":
            x = to_number(self.scalar(self.ev(node[1], sheet)))
            return x if isinstance(x, XLError) else x / 100
        if kind == "invoke":
            # The callee is evaluated first, so LAMBDA(...)(...) is refused with the same message as
            # a plain LAMBDA(...). Any other callee gives a value, which cannot be called.
            self.ev(node[1], sheet)
            raise Unsupported("unexpected text after the formula")
        if kind == "bin":
            return self.binary(node[1], node[2], node[3], sheet)
        if kind == "func":
            return self.call(node[1], node[2], sheet)
        if kind == "name":
            upper = node[1].upper()
            for frame in reversed(self.scopes):
                if upper in frame:
                    return frame[upper]
            return self.resolve_name(node[1], sheet)
        raise Unsupported(f"unknown expression kind {kind}")

    def binary(self, op, left, right, sheet):
        a = self.ev(left, sheet)
        b = self.ev(right, sheet)
        if self._is_array(a) or self._is_array(b):
            self.array_used = True
            return self._elementwise(op, a, b)
        return self._apply(op, self.scalar(a), self.scalar(b))

    @staticmethod
    def _is_array(v):
        """True for a range or array with more than one cell."""
        if isinstance(v, (Ref, Arr)):
            return v.height * v.width > 1
        return False

    def _elementwise(self, op, a, b):
        """An operator applied cell by cell: to two arrays of the same size, or to an array and a
        single value."""
        if self._is_array(a) and self._is_array(b):
            ga, gb = self._as_arr(a), self._as_arr(b)
            if ga.height != gb.height or ga.width != gb.width:
                raise Unsupported("the operands are arrays of different sizes")
            return Arr([[self._apply(op, x, y) for x, y in zip(la, lb)]
                        for la, lb in zip(ga.rows, gb.rows)])
        if self._is_array(a):
            sb = self.scalar(b)
            return Arr([[self._apply(op, x, sb) for x in line] for line in self._as_arr(a).rows])
        sa = self.scalar(a)
        return Arr([[self._apply(op, sa, y) for y in line] for line in self._as_arr(b).rows])

    def _apply(self, op, a, b):
        """One operator on two single values."""
        if isinstance(a, XLError):
            return a
        if isinstance(b, XLError):
            return b
        if op in ("=", "<>", "<", ">", "<=", ">="):
            return compare_values(op, a, b)
        if op == "&":
            return to_text(a) + to_text(b)
        x = to_number(a)
        if isinstance(x, XLError):
            return x
        y = to_number(b)
        if isinstance(y, XLError):
            return y
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            return DIV0 if y == 0 else x / y
        if op == "^":
            try:
                result = x ** y
            except ZeroDivisionError:
                return DIV0
            except OverflowError:
                return NUM
            return NUM if isinstance(result, complex) else result
        raise Unsupported(f"operator {op}")

    def call(self, name, nodes, sheet):
        if name in ARRAY_RESULT_FUNCTIONS:
            self.array_used = True
        if name == "IF":
            return self.fn_IF(nodes, sheet)
        if name == "IFERROR":
            return self.fn_IFERROR(nodes, sheet)
        if name == "LET":
            return self.eval_let(nodes, sheet)
        handler = getattr(self, "fn_" + name, None)
        if handler is None:
            raise Unsupported(f"the function {name} is not supported")
        raws = [self.ev(arg, sheet) for arg in nodes]
        if name not in ARRAY_ARGUMENT_FUNCTIONS:
            raws = [self.single_arg(raw, name) for raw in raws]
        return handler(raws)

    @staticmethod
    def single_arg(raw, name):
        """An array is accepted only by the functions in ARRAY_ARGUMENT_FUNCTIONS."""
        if isinstance(raw, Arr):
            if raw.height == 1 and raw.width == 1:
                return raw.rows[0][0]
            raise Unsupported(f"an array result is not accepted as an argument of {name} in this version")
        return raw

    def eval_let(self, nodes, sheet):
        """LET(name, value, [name, value, ...], calculation). Names are visible only inside the call."""
        if len(nodes) < 3 or len(nodes) % 2 == 0:
            raise Unsupported("LET needs one or more name and value pairs followed by a calculation")
        frame = {}
        self.scopes.append(frame)
        try:
            for i in range(0, len(nodes) - 1, 2):
                name_node = nodes[i]
                if name_node[0] == "ref" and name_node[2] == name_node[4] and name_node[3] == name_node[5]:
                    raise Unsupported(f'"{coord(name_node[2], name_node[3])}" is not a valid LET name')
                if name_node[0] != "name":
                    raise Unsupported("a LET name must be a plain name, not a value or a reference")
                name = name_node[1]
                if not valid_let_name(name):
                    raise Unsupported(f'"{name}" is not a valid LET name')
                if name.upper() in frame:
                    raise Unsupported(f'the LET name "{name}" is given more than once')
                frame[name.upper()] = self.ev(nodes[i + 1], sheet)
            return self.ev(nodes[-1], sheet)
        finally:
            self.scopes.pop()

    def _flat(self, raw):
        """The values of a range or array, row by row; a single value gives one item."""
        if isinstance(raw, Ref):
            return self.iter_ref(raw)
        if isinstance(raw, Arr):
            return [v for line in raw.rows for v in line]
        return [raw]

    def _as_arr(self, raw):
        if isinstance(raw, Arr):
            return raw
        if isinstance(raw, Ref):
            return Arr([[self.cell_value(raw.sheet, r, c) for c in range(raw.c1, raw.c2 + 1)]
                        for r in range(raw.r1, raw.r2 + 1)])
        return Arr([[raw]])

    def _opt(self, raws, pos, default):
        """An optional scalar argument; an omitted or empty argument gives the default."""
        if len(raws) <= pos:
            return default
        v = self.scalar(raws[pos])
        return default if v is None else v

    # -- lazy functions
    def fn_IF(self, nodes, sheet):
        if not 1 <= len(nodes) <= 3:
            raise Unsupported("IF needs 2 or 3 arguments")
        cond = to_bool(self.scalar(self.ev(nodes[0], sheet)))
        if isinstance(cond, XLError):
            return cond
        if cond:
            return self.ev(nodes[1], sheet) if len(nodes) > 1 else True
        return self.ev(nodes[2], sheet) if len(nodes) > 2 else False

    def fn_IFERROR(self, nodes, sheet):
        if len(nodes) != 2:
            raise Unsupported("IFERROR needs 2 arguments")
        value = self.scalar(self.ev(nodes[0], sheet))
        if isinstance(value, XLError):
            return self.ev(nodes[1], sheet)
        return value

    # -- helpers for argument lists
    def _numbers(self, raws):
        out = []
        for raw in raws:
            if isinstance(raw, (Ref, Arr)):
                for v in self._flat(raw):
                    if isinstance(v, XLError):
                        return v
                    if is_number(v):
                        out.append(v)
            else:
                if raw is None:
                    continue
                if isinstance(raw, XLError):
                    return raw
                n = to_number(raw)
                if isinstance(n, XLError):
                    return n
                out.append(n)
        return out

    def _criteria_pairs(self, raws, start):
        if len(raws) <= start or (len(raws) - start) % 2:
            return VALUE
        pairs = []
        for i in range(start, len(raws), 2):
            rng = raws[i]
            crit = self.scalar(raws[i + 1])
            if not isinstance(rng, Ref):
                return VALUE
            if isinstance(crit, XLError):
                return crit
            pairs.append((rng, make_predicate(crit)))
        return pairs

    @staticmethod
    def _same_shape(refs):
        base = refs[0]
        return all(r.height == base.height and r.width == base.width for r in refs)

    def _matching_positions(self, pairs):
        base = pairs[0][0]
        for i in range(base.height):
            for j in range(base.width):
                if all(pred(self.cell_value(rng.sheet, rng.r1 + i, rng.c1 + j)) for rng, pred in pairs):
                    yield i, j

    def _ifs_values(self, raws):
        """Numeric values from the first range where every criterion pair matches."""
        target = raws[0]
        if not isinstance(target, Ref):
            return VALUE
        pairs = self._criteria_pairs(raws, 1)
        if isinstance(pairs, XLError):
            return pairs
        if not self._same_shape([target] + [p[0] for p in pairs]):
            return VALUE
        values = []
        for i, j in self._matching_positions(pairs):
            v = self.cell_value(target.sheet, target.r1 + i, target.c1 + j)
            if isinstance(v, XLError):
                return v
            if is_number(v):
                values.append(v)
        return values

    def _approx_index(self, items, key, direction):
        """Approximate MATCH and VLOOKUP by binary search.

        direction 1: largest value <= key in ascending data. direction -1: smallest value >= key
        in descending data. Blanks and values of another type than the key are ignored."""
        rank = _rank(key)
        pairs = [(i, v) for i, v in enumerate(items) if v is not None and _rank(v) == rank]
        pairs = _ordered_pairs(pairs, direction)
        keys = [_order_key(v) for _, v in pairs]
        mode = "largest_le" if direction == 1 else "smallest_ge"
        pos = _bsearch_pick(keys, _order_key(key), mode)
        return None if pos is None else pairs[pos][0]

    @staticmethod
    def _exact_index(items, key, wildcard):
        pattern = wildcard_regex(key) if wildcard and isinstance(key, str) and _has_wildcard(key) else None
        for i, v in enumerate(items):
            if v is None or isinstance(v, XLError):
                continue
            if pattern is not None:
                if isinstance(v, str) and pattern.fullmatch(v):
                    return i
                continue
            if _loose_eq(key, v):
                return i
        return None

    # -- functions: aggregates
    def fn_SUM(self, raws):
        items = self._numbers(raws)
        return items if isinstance(items, XLError) else sum(items)

    def fn_AVERAGE(self, raws):
        items = self._numbers(raws)
        if isinstance(items, XLError):
            return items
        return DIV0 if not items else sum(items) / len(items)

    def fn_MIN(self, raws):
        items = self._numbers(raws)
        if isinstance(items, XLError):
            return items
        return min(items) if items else 0

    def fn_MAX(self, raws):
        items = self._numbers(raws)
        if isinstance(items, XLError):
            return items
        return max(items) if items else 0

    def fn_COUNT(self, raws):
        count = 0
        for raw in raws:
            if isinstance(raw, (Ref, Arr)):
                count += sum(1 for v in self._flat(raw) if is_number(v))
            elif raw is None or isinstance(raw, XLError):
                continue
            elif isinstance(raw, bool) or is_number(raw):
                count += 1
            elif isinstance(raw, str) and _NUM_TEXT_RE.fullmatch(raw.strip()):
                count += 1
        return count

    def fn_COUNTA(self, raws):
        count = 0
        for raw in raws:
            if isinstance(raw, (Ref, Arr)):
                count += sum(1 for v in self._flat(raw) if v is not None)
            elif raw is not None:
                count += 1
        return count

    def fn_TEXTJOIN(self, raws):
        if len(raws) < 3:
            return VALUE
        delimiter = self.scalar(raws[0])
        if isinstance(delimiter, XLError):
            return delimiter
        ignore_empty = to_bool(self.scalar(raws[1]))
        if isinstance(ignore_empty, XLError):
            return ignore_empty
        parts = []
        for raw in raws[2:]:
            values = self._flat(raw)
            for v in values:
                if isinstance(v, XLError):
                    return v
                text = to_text(v)
                if ignore_empty and text == "":
                    continue
                parts.append(text)
        return to_text(delimiter).join(parts)

    # -- functions: conditional aggregates
    def fn_COUNTIF(self, raws):
        if len(raws) != 2:
            return VALUE
        return self.fn_COUNTIFS(raws)

    def fn_COUNTIFS(self, raws):
        pairs = self._criteria_pairs(raws, 0)
        if isinstance(pairs, XLError):
            return pairs
        if not self._same_shape([p[0] for p in pairs]):
            return VALUE
        return sum(1 for _ in self._matching_positions(pairs))

    def fn_SUMIF(self, raws):
        if len(raws) not in (2, 3):
            return VALUE
        rng = raws[0]
        if not isinstance(rng, Ref):
            return VALUE
        crit = self.scalar(raws[1])
        if isinstance(crit, XLError):
            return crit
        sum_rng = raws[2] if len(raws) == 3 else rng
        if not isinstance(sum_rng, Ref):
            return VALUE
        pred = make_predicate(crit)
        total = 0
        for i in range(rng.height):
            for j in range(rng.width):
                if pred(self.cell_value(rng.sheet, rng.r1 + i, rng.c1 + j)):
                    v = self.cell_value(sum_rng.sheet, sum_rng.r1 + i, sum_rng.c1 + j)
                    if isinstance(v, XLError):
                        return v
                    if is_number(v):
                        total += v
        return total

    def fn_SUMIFS(self, raws):
        values = self._ifs_values(raws)
        return values if isinstance(values, XLError) else sum(values)

    def fn_AVERAGEIFS(self, raws):
        values = self._ifs_values(raws)
        if isinstance(values, XLError):
            return values
        return DIV0 if not values else sum(values) / len(values)

    def fn_MAXIFS(self, raws):
        values = self._ifs_values(raws)
        if isinstance(values, XLError):
            return values
        return max(values) if values else 0

    def fn_MINIFS(self, raws):
        values = self._ifs_values(raws)
        if isinstance(values, XLError):
            return values
        return min(values) if values else 0

    # -- functions: lookup and reference
    def _vh_lookup(self, raws, vertical):
        if len(raws) not in (3, 4):
            return VALUE
        key = self.scalar(raws[0])
        if isinstance(key, XLError):
            return key
        table = raws[1]
        if not isinstance(table, Ref):
            return VALUE
        index = to_number(self.scalar(raws[2]))
        if isinstance(index, XLError):
            return index
        approx = True
        if len(raws) == 4:
            approx = to_bool(self.scalar(raws[3]))
            if isinstance(approx, XLError):
                return approx
        index = int(index)
        limit = table.width if vertical else table.height
        if index < 1:
            return VALUE
        if index > limit:
            return REF
        if vertical:
            items = [self.cell_value(table.sheet, r, table.c1) for r in range(table.r1, table.r2 + 1)]
        else:
            items = [self.cell_value(table.sheet, table.r1, c) for c in range(table.c1, table.c2 + 1)]
        key = _lookup_key(key)
        pos = self._approx_index(items, key, 1) if approx else self._exact_index(items, key, True)
        if pos is None:
            return NA
        if vertical:
            return self.cell_value(table.sheet, table.r1 + pos, table.c1 + index - 1)
        return self.cell_value(table.sheet, table.r1 + index - 1, table.c1 + pos)

    def fn_VLOOKUP(self, raws):
        return self._vh_lookup(raws, vertical=True)

    def fn_HLOOKUP(self, raws):
        return self._vh_lookup(raws, vertical=False)

    def fn_MATCH(self, raws):
        if len(raws) not in (2, 3):
            return VALUE
        key = self.scalar(raws[0])
        if isinstance(key, XLError):
            return key
        arr = raws[1]
        if not isinstance(arr, Ref) or (arr.height != 1 and arr.width != 1):
            return VALUE
        mode = 1
        if len(raws) == 3:
            mode = to_number(self.scalar(raws[2]))
            if isinstance(mode, XLError):
                return mode
            mode = int(mode)
        items = list(self.iter_ref(arr))
        key = _lookup_key(key)
        if mode == 0:
            pos = self._exact_index(items, key, True)
        elif mode in (1, -1):
            pos = self._approx_index(items, key, mode)
        else:
            raise Unsupported(f"MATCH with match_type {mode} is not implemented (only 0, 1 and -1)")
        return NA if pos is None else pos + 1

    def fn_INDEX(self, raws):
        if len(raws) not in (2, 3):
            return VALUE
        arr = raws[0]
        if not isinstance(arr, Ref):
            return VALUE
        row = to_number(self.scalar(raws[1]))
        if isinstance(row, XLError):
            return row
        col = None
        if len(raws) == 3:
            col = to_number(self.scalar(raws[2]))
            if isinstance(col, XLError):
                return col
            col = int(col)
        row = int(row)
        if col is None:
            if arr.height == 1 and arr.width > 1:
                row, col = 1, row
            elif arr.width == 1:
                col = 1
            else:
                raise Unsupported("INDEX without a column on a two-dimensional range returns a range")
        if row == 0 or col == 0:
            raise Unsupported("INDEX with row or column 0 returns a range")
        if row < 1 or row > arr.height or col < 1 or col > arr.width:
            return REF
        return self.cell_value(arr.sheet, arr.r1 + row - 1, arr.c1 + col - 1)

    def _xlookup_binary(self, items, key, match_mode, search_mode):
        """XLOOKUP search_mode 2 (ascending data) or -2 (descending data): binary search.
        Returns the position in items, or None."""
        if match_mode == 2:
            raise Unsupported("XLOOKUP wildcard matching (match_mode 2) with a binary search mode "
                              "is not implemented")
        pairs = [(i, v) for i, v in enumerate(items) if v is not None and _rank(v) == _rank(key)]
        pairs = _ordered_pairs(pairs, 1 if search_mode == 2 else -1)
        keys = [_order_key(v) for _, v in pairs]
        mode = {0: "exact", -1: "next_smaller", 1: "next_larger"}[match_mode]
        pos = _bsearch_pick(keys, _order_key(key), mode)
        return None if pos is None else pairs[pos][0]

    def fn_XLOOKUP(self, raws):
        if not 3 <= len(raws) <= 6:
            return VALUE
        key = self.scalar(raws[0])
        if isinstance(key, XLError):
            return key
        look, ret = raws[1], raws[2]
        if not (isinstance(look, Ref) and isinstance(ret, Ref)):
            return VALUE
        if look.height != 1 and look.width != 1:
            return VALUE
        # A vertical lookup range (one column) takes a return range with the same number of rows,
        # and returns the matching row across it. A horizontal one works the same with columns.
        vertical = look.width == 1
        n = look.height * look.width
        if (ret.height if vertical else ret.width) != n:
            return VALUE
        has_default = len(raws) > 3
        if_not_found = raws[3] if has_default else None
        match_mode = 0
        search_mode = 1
        if len(raws) > 4:
            match_mode = to_number(self.scalar(raws[4]))
            if isinstance(match_mode, XLError):
                return match_mode
            match_mode = int(match_mode)
        if len(raws) > 5:
            search_mode = to_number(self.scalar(raws[5]))
            if isinstance(search_mode, XLError):
                return search_mode
            search_mode = int(search_mode)
        if search_mode not in (1, -1, 2, -2) or match_mode not in (-1, 0, 1, 2):
            return VALUE
        items = list(self.iter_ref(look))
        key = _lookup_key(key)
        found = None
        if search_mode in (2, -2):
            found = self._xlookup_binary(items, key, match_mode, search_mode)
            order = ()  # binary modes do not scan; the approximate scan below then finds nothing
        else:
            order = range(n) if search_mode == 1 else range(n - 1, -1, -1)
        pattern = wildcard_regex(key) if match_mode == 2 and isinstance(key, str) else None
        for idx in order:
            v = items[idx]
            if v is None or isinstance(v, XLError):
                continue
            if pattern is not None:
                hit = isinstance(v, str) and pattern.fullmatch(v) is not None
            else:
                hit = _loose_eq(key, v)
            if hit:
                found = idx
                break
        if found is None and match_mode in (-1, 1):
            rank = _rank(key)
            best = None
            for idx in order:
                v = items[idx]
                if v is None or _rank(v) != rank:
                    continue
                kv, kk = _order_key(v), _order_key(key)
                if match_mode == -1 and kv < kk and (best is None or kv > _order_key(items[best])):
                    best = idx
                if match_mode == 1 and kv > kk and (best is None or kv < _order_key(items[best])):
                    best = idx
            found = best
        if found is None:
            if has_default:
                if if_not_found is None:
                    return 0
                return self.scalar(if_not_found)
            return NA
        if vertical:
            row = ret.r1 + found
            if ret.width == 1:
                return self.cell_value(ret.sheet, row, ret.c1)
            self.array_used = True
            return Arr([[self.cell_value(ret.sheet, row, c) for c in range(ret.c1, ret.c2 + 1)]])
        col = ret.c1 + found
        if ret.height == 1:
            return self.cell_value(ret.sheet, ret.r1, col)
        self.array_used = True
        return Arr([[self.cell_value(ret.sheet, r, col)] for r in range(ret.r1, ret.r2 + 1)])

    # -- functions: dynamic arrays (FILTER, SORT, UNIQUE, SEQUENCE)
    def fn_FILTER(self, raws):
        if not 2 <= len(raws) <= 3:
            return VALUE
        for raw in raws[:2]:
            if isinstance(raw, XLError):
                return raw
        arr = self._as_arr(raws[0])
        inc = self._as_arr(raws[1])
        if inc.width == 1 and inc.height == arr.height:
            flags = [line[0] for line in inc.rows]
            by_row = True
        elif inc.height == 1 and inc.width == arr.width:
            flags = inc.rows[0]
            by_row = False
        else:
            return VALUE
        keep = []
        for flag in flags:
            truth = _include_flag(flag)
            if isinstance(truth, XLError):
                return truth
            keep.append(truth)
        if by_row:
            picked = [line for line, k in zip(arr.rows, keep) if k]
            return Arr(picked) if picked else self._filter_empty(raws)
        cols = [j for j, k in enumerate(keep) if k]
        if not cols:
            return self._filter_empty(raws)
        return Arr([[line[j] for j in cols] for line in arr.rows])

    def _filter_empty(self, raws):
        """FILTER with no matching rows: the if_empty argument, or #CALC! when it is omitted."""
        if len(raws) < 3:
            return CALC
        fallback = raws[2]
        if isinstance(fallback, Ref):
            return self._as_arr(fallback)
        return fallback

    def fn_SORT(self, raws):
        if not 1 <= len(raws) <= 4:
            return VALUE
        if isinstance(raws[0], XLError):
            return raws[0]
        arr = self._as_arr(raws[0])
        index = to_number(self._opt(raws, 1, 1))
        if isinstance(index, XLError):
            return index
        order = to_number(self._opt(raws, 2, 1))
        if isinstance(order, XLError):
            return order
        by_col = to_bool(self._opt(raws, 3, False))
        if isinstance(by_col, XLError):
            return by_col
        if order not in (1, -1):
            return VALUE
        lines = arr.transposed().rows if by_col else arr.rows
        index = int(index)
        if index < 1 or index > len(lines[0]):
            return VALUE
        keys = [line[index - 1] for line in lines]
        for key in keys:
            if isinstance(key, XLError):
                return key
        ordered = _sort_indices(keys, descending=(order == -1))
        result = Arr([lines[i] for i in ordered])
        return result.transposed() if by_col else result

    def fn_UNIQUE(self, raws):
        if not 1 <= len(raws) <= 3:
            return VALUE
        if isinstance(raws[0], XLError):
            return raws[0]
        arr = self._as_arr(raws[0])
        by_col = to_bool(self._opt(raws, 1, False))
        if isinstance(by_col, XLError):
            return by_col
        once = to_bool(self._opt(raws, 2, False))
        if isinstance(once, XLError):
            return once
        lines = arr.transposed().rows if by_col else arr.rows
        counts = {}
        firsts = {}
        order = []
        for line in lines:
            key = tuple(_uniq_key(v) for v in line)
            if key in counts:
                counts[key] += 1
            else:
                counts[key] = 1
                firsts[key] = line
                order.append(key)
        chosen = [firsts[k] for k in order if counts[k] == 1 or not once]
        if not chosen:
            return CALC
        result = Arr([list(line) for line in chosen])
        return result.transposed() if by_col else result

    def fn_SEQUENCE(self, raws):
        if not 1 <= len(raws) <= 4:
            return VALUE
        first = self._opt(raws, 0, None)
        if first is None:
            return VALUE
        rows = to_number(first)
        cols = to_number(self._opt(raws, 1, 1))
        start = to_number(self._opt(raws, 2, 1))
        step = to_number(self._opt(raws, 3, 1))
        for v in (rows, cols, start, step):
            if isinstance(v, XLError):
                return v
        rows, cols = int(rows), int(cols)
        if rows < 1 or cols < 1:
            return CALC
        if rows * cols > 1000000:
            raise Unsupported("the SEQUENCE result is too large for the built-in evaluator")
        return Arr([[start + step * (i * cols + j) for j in range(cols)] for i in range(rows)])


def _include_flag(v):
    """FILTER include value: TRUE or a non-zero number keeps the row; FALSE, zero and blank drop it;
    text gives #VALUE!; an error is passed on."""
    if isinstance(v, XLError):
        return v
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if is_number(v):
        return v != 0
    return VALUE


def _sort_rank(v):
    """Ascending order for non-blank values: numbers, then text (case-insensitive), then TRUE/FALSE."""
    if isinstance(v, bool):
        return (2, v)
    if is_number(v):
        return (0, v)
    return (1, v.casefold())


def _sort_indices(keys, descending):
    """Positions of the keys in sorted order. Blanks go last in both directions; equal keys keep
    their original order."""
    present = [i for i, k in enumerate(keys) if k is not None]
    blanks = [i for i, k in enumerate(keys) if k is None]
    present.sort(key=lambda i: _sort_rank(keys[i]), reverse=descending)
    return present + blanks


def _uniq_key(v):
    """Equality key for UNIQUE: text ignores case, numbers compare by value, blanks are their own value."""
    if v is None:
        return ("blank",)
    if isinstance(v, bool):
        return ("bool", v)
    if is_number(v):
        return ("num", float(v))
    if isinstance(v, str):
        return ("text", v.casefold())
    return ("err", v.code)


def collect_names(wb):
    """Map (scope, NAME) to the stored text of each defined name.

    scope is None for workbook-level names and the sheet title for sheet-level names."""
    table = {}
    for name, dn in wb.defined_names.items():
        table[(None, name.upper())] = dn.attr_text or ""
    for ws in wb.worksheets:
        for name, dn in ws.defined_names.items():
            table[(ws.title, name.upper())] = dn.attr_text or ""
    return table


class SheetModel:
    __slots__ = ("cells", "max_row")

    def __init__(self, cells, max_row):
        self.cells = cells
        self.max_row = max_row


_ARRAY_CALL_RE = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.|_xlws\.)*(?:FILTER|SORT|UNIQUE|SEQUENCE|LET)\s*\(",
                            re.IGNORECASE)


def _block_cells(anchor_cell):
    """The cells of an array formula's block other than its anchor, or an empty list."""
    ref = anchor_cell.value.ref
    if ":" not in ref:
        return []
    r1, c1, r2, c2 = parse_range_text(ref)
    return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
            if (r, c) != (anchor_cell.row, anchor_cell.column)]


def array_candidates(wb):
    """Formula cells that may return an array, per sheet (sorted by row, then column), and the
    anchors of multi-cell array formulas that are not array functions."""
    candidates = {}
    anchors = set()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f":
                    continue
                multi = isinstance(cell.value, ArrayFormula) and ":" in cell.value.ref
                text = formula_text(cell)
                is_array_call = bool(_ARRAY_CALL_RE.search(text))
                if multi and not is_array_call:
                    anchors.add((ws.title, cell.row, cell.column))
                if multi or is_array_call:
                    candidates.setdefault(ws.title, []).append((cell.row, cell.column))
    for lst in candidates.values():
        lst.sort()
    return candidates, anchors


def build_model(wb):
    model = {}
    for ws in wb.worksheets:
        # cells inside an existing array block belong to its formula, so they have no constant value
        owned = set()
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, ArrayFormula):
                    owned.update(_block_cells(cell))
        cells = {}
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                key = (cell.row, cell.column)
                if key in owned:
                    continue
                if cell.data_type == "f":
                    text = v.text if isinstance(v, ArrayFormula) else v
                    cells[key] = ("f", text[1:] if text.startswith("=") else text)
                elif isinstance(v, (dt.datetime, dt.date, dt.time)):
                    cells[key] = ("v", to_excel(v))
                elif cell.data_type == "e":
                    cells[key] = ("v", XLError(str(v)))
                else:
                    cells[key] = ("v", v)
        model[ws.title] = SheetModel(cells, ws.max_row)
    return model


def formula_text(cell):
    v = cell.value
    if isinstance(v, ArrayFormula):
        return "=" + v.text.lstrip("=")
    return v


def clear_array_members(wb):
    """Empty every cell inside an existing array block. The anchor keeps its formula. Cells that
    recalc writes again get their values back; the rest stay empty, as saving drops stored results."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, ArrayFormula):
                    for r, c in _block_cells(cell):
                        ws.cell(r, c).value = None


def set_spill_value(cell, value):
    """Write one cached value of a spilled block as a constant, the way Excel stores block cells."""
    if isinstance(value, XLError):
        cell.value = value.code
        cell.data_type = "s" if value.code in TEXT_STORED_ERRORS else "e"
        return
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"


def _cache_xml(value):
    if isinstance(value, XLError):
        if value.code in TEXT_STORED_ERRORS:
            return "str", value.code
        return "e", value.code
    if isinstance(value, bool):
        return "b", "1" if value else "0"
    if isinstance(value, str):
        return "str", value
    if isinstance(value, float) and value.is_integer() and abs(value) < 1e15:
        return None, str(int(value))
    if isinstance(value, float):
        return None, repr(value)
    return None, str(value)


def _sheet_parts(zf):
    wb_root = etree.fromstring(zf.read("xl/workbook.xml"), SAFE_PARSER)
    rel_root = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"), SAFE_PARSER)
    targets = {rel.get("Id"): rel.get("Target") for rel in rel_root}
    parts = {}
    for sheet in wb_root.iter(f"{{{MAIN_NS}}}sheet"):
        target = targets.get(sheet.get(f"{{{OFFICE_REL_NS}}}id"), "")
        if target.startswith("/"):
            part = target[1:]
        else:
            part = posixpath.normpath(posixpath.join("xl", target))
        parts[part] = sheet.get("name")
    return parts


# The dynamic-array property that Excel attaches to a formula cell (cell metadata 1 = XLDAPR).
# Excel writes this part itself; the same minimal form is written here.
DYNAMIC_METADATA_PART = "xl/metadata.xml"
DYNAMIC_METADATA = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    b'<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" '
    b'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    b'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    b'<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    b'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" '
    b'fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
    b'<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)
PKG_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _patch_sheet_xml(data, cache, dynamic=frozenset()):
    root = etree.fromstring(data, SAFE_PARSER)
    for c in root.iter(f"{{{MAIN_NS}}}c"):
        if c.find(f"{{{MAIN_NS}}}f") is None:
            continue
        ref = c.get("r")
        if ref in dynamic:
            c.set("cm", "1")
        if ref not in cache:
            continue
        t, text = _cache_xml(cache[ref])
        v = c.find(f"{{{MAIN_NS}}}v")
        if v is None:
            v = etree.SubElement(c, f"{{{MAIN_NS}}}v")
        v.text = text
        if t is None:
            c.attrib.pop("t", None)
        else:
            c.set("t", t)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_dynamic_metadata_parts(data):
    """Register xl/metadata.xml in the content types and the workbook relationships."""
    ct = etree.fromstring(data["[Content_Types].xml"], SAFE_PARSER)
    if not any(o.get("PartName") == "/" + DYNAMIC_METADATA_PART for o in ct):
        override = etree.SubElement(ct, f"{{{PKG_CT_NS}}}Override")
        override.set("PartName", "/" + DYNAMIC_METADATA_PART)
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")
    data["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    rels = etree.fromstring(data["xl/_rels/workbook.xml.rels"], SAFE_PARSER)
    used = {rel.get("Id") for rel in rels}
    rel_id = "rIdDynArr1"
    n = 1
    while rel_id in used:
        n += 1
        rel_id = f"rIdDynArr{n}"
    rel = etree.SubElement(rels, f"{{{PKG_REL_NS}}}Relationship")
    rel.set("Id", rel_id)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata")
    rel.set("Target", "metadata.xml")
    data["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8",
                                                        standalone=True)
    data[DYNAMIC_METADATA_PART] = DYNAMIC_METADATA


def inject_cached_values(xlsx_bytes, dest, cache_by_sheet, dynamic_by_sheet=None):
    """Copy the workbook, adding cached <v> values to formula cells. Cells in dynamic_by_sheet are
    marked as dynamic-array formulas, and the metadata part that Excel needs for that is added."""
    dynamic_by_sheet = dynamic_by_sheet or {}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zin:
        parts = _sheet_parts(zin)
        names = [info.filename for info in zin.infolist()]
        data = {name: zin.read(name) for name in names}
    for name, sheet in parts.items():
        data[name] = _patch_sheet_xml(data[name], cache_by_sheet.get(sheet, {}),
                                      dynamic_by_sheet.get(sheet, frozenset()))
    if any(dynamic_by_sheet.values()):
        _add_dynamic_metadata_parts(data)
        if DYNAMIC_METADATA_PART not in names:
            names.append(DYNAMIC_METADATA_PART)
    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, data[name])


def _recalc_error_message(exc, fallback):
    where = exc.where
    loc = f"{where[0]}!{coord(where[1], where[2])}" if where else fallback
    return (f"Formula in {loc} is not supported by the built-in evaluator: {exc.reason}. "
            f"Nothing was written.")


# ---------------------------------------------------------------- commands

FORMULA_CACHE_NOTE = ("Saving with this tool drops stored formula results for every formula in the "
                      "workbook. Run recalc to store them again.")


def count_formulas(wb):
    return sum(1 for ws in wb.worksheets for row in ws.iter_rows() for cell in row
               if cell.data_type == "f")


def cmd_sheets(args):
    wb = open_workbook(args.file)
    sheets = [{"name": ws.title, "state": ws.sheet_state, "dimensions": ws.dimensions,
               "max_row": ws.max_row, "max_column": ws.max_column} for ws in wb.worksheets]
    lines = [f"{s['name']}\t{s['state']}\t{s['dimensions']}\t{s['max_row']}\t{s['max_column']}"
             for s in sheets]
    emit(args, {"sheets": sheets}, lines)
    return 0


def cmd_read(args):
    wb_f = open_workbook(args.file)
    ws = pick_sheet(wb_f, args.sheet)
    if args.range:
        r1, c1, r2, c2 = parse_range_text(args.range)
    else:
        r1, c1, r2, c2 = used_range(ws)
    wb_v = open_workbook(args.file, data_only=True) if args.mode == "values" else None
    wsv = wb_v[ws.title] if wb_v is not None else None
    rows = []
    not_calculated = []
    for r in range(r1, r2 + 1):
        row = []
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            if args.mode == "formulas":
                row.append(json_value(formula_text(cell)))
            else:
                value = wsv.cell(r, c).value
                if cell.data_type == "f" and value is None:
                    not_calculated.append(coord(r, c))
                row.append(json_value(value))
        rows.append(row)
    rng = f"{coord(r1, c1)}:{coord(r2, c2)}"
    payload = {"sheet": ws.title, "range": rng, "rows": rows, "not_calculated": not_calculated}
    lines = ["\t".join("" if v is None else str(v) for v in row) for row in rows]
    if args.out:
        target = plan_target(args.file, args.out, False, None, args.force, require_xlsx=False)
        write_text_output(target, json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    emit(args, payload, lines)
    if not_calculated:
        warn(f"{len(not_calculated)} formula(s) have no calculated value. Run recalc.")
    return 0


def convert_formula_cell(entry, n, ref):
    text = entry["formula"]
    if not text.startswith("="):
        raise ToolError(f'Formula in {ref} must start with "=".')
    return text


def cmd_write(args):
    src = args.file
    check_input(src)
    target = plan_target(src, args.out, args.in_place, "write", args.force)
    entries = load_patch(args.patch)
    allowed = drawing_guard(src, args.allow_drawing_rewrite)
    wb = open_workbook(src)
    clear_array_members(wb)
    ops = []
    prefixed = []
    for n, entry in enumerate(entries, 1):
        ws = pick_sheet(wb, entry["sheet"])
        row, col = parse_cell_text(entry["cell"])
        ref = coord(row, col)
        cell = ws.cell(row, col)
        if "formula" in entry:
            text = convert_formula_cell(entry, n, ref)
            text, names = add_xlfn_prefix(text)
            for name in names:
                prefixed.append({"cell": f"{ws.title}!{ref}", "function": name})
            cell.value = text
            expected = text
            is_formula = True
        else:
            expected = convert_patch_value(entry, n, ref)
            cell.value = expected
            is_formula = False
        if entry.get("number_format") is not None:
            cell.number_format = entry["number_format"]
        elif entry.get("type") == "date":
            cell.number_format = DATE_PRESET
        ops.append((ws.title, ref, expected, is_formula))

    def save(path):
        wb.save(str(path))

    def verify(path):
        wb2 = open_workbook(path)
        for sheet, ref, expected, _ in ops:
            got = wb2[sheet][ref].value
            if not values_equal(expected, got):
                raise ToolError(f"{sheet}!{ref} reads back as {got!r}, expected {expected!r}")

    total_formulas = count_formulas(wb)
    commit_workbook(save, target, verify)
    payload = {"output": str(target), "cells_written": len(ops), "formula_cells": total_formulas,
               "prefixed_functions": prefixed, "drawings_rewritten": allowed}
    lines = [f"Wrote {target} ({len(ops)} cell(s))."]
    lines += [f"Prefixed {p['function']} in {p['cell']} with _xlfn." for p in prefixed]
    if total_formulas:
        payload["note"] = FORMULA_CACHE_NOTE
        lines.append(FORMULA_CACHE_NOTE)
    emit(args, payload, lines)
    return 0


def cmd_format(args):
    src = args.file
    check_input(src)
    target = plan_target(src, args.out, args.in_place, "format", args.force)
    if (args.preset is None) == (args.code is None):
        raise ToolError("Give exactly one of --preset or --code.")
    if args.preset is not None:
        if args.preset == "currency":
            if '"' in args.currency_symbol:
                raise ToolError("--currency-symbol must not contain a double quote.")
            code = f'"{args.currency_symbol}"#,##0.00'
        elif args.preset == "date":
            code = DATE_PRESET
        else:
            code = PERCENT_PRESET
    else:
        code = args.code
        if not code.strip():
            raise ToolError("The number format code must not be empty.")
    allowed = drawing_guard(src, args.allow_drawing_rewrite)
    wb = open_workbook(src)
    clear_array_members(wb)
    ws = pick_sheet(wb, args.sheet)
    if args.range:
        r1, c1, r2, c2 = parse_range_text(args.range)
    else:
        r1, c1, r2, c2 = used_range(ws)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).number_format = code
    sheet_title = ws.title
    rng = f"{coord(r1, c1)}:{coord(r2, c2)}"

    def save(path):
        wb.save(str(path))

    def verify(path):
        wb2 = open_workbook(path)
        ws2 = wb2[sheet_title]
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                got = ws2.cell(r, c).number_format
                if got != code:
                    raise ToolError(f"{coord(r, c)} has number format {got!r}, expected {code!r}")

    total_formulas = count_formulas(wb)
    commit_workbook(save, target, verify)
    cells = (r2 - r1 + 1) * (c2 - c1 + 1)
    payload = {"output": str(target), "sheet": sheet_title, "range": rng,
               "number_format": code, "cells_formatted": cells, "formula_cells": total_formulas,
               "drawings_rewritten": allowed}
    lines = [f"Formatted {rng} on {sheet_title} as {code} ({cells} cell(s)). Wrote {target}."]
    if total_formulas:
        payload["note"] = FORMULA_CACHE_NOTE
        lines.append(FORMULA_CACHE_NOTE)
    emit(args, payload, lines)
    return 0


def _recalc_scope_cells(cells_arg, ws_title):
    out = []
    for part in cells_arg.split(","):
        if not part.strip():
            raise ToolError(f'"{cells_arg}" is not a valid cell or range.')
        r1, c1, r2, c2 = parse_range_text(part)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                out.append((ws_title, r, c))
    return out


def cmd_recalc(args):
    src = args.file
    check_input(src)
    target = plan_target(src, args.out, args.in_place, "recalc", args.force)
    if args.timeout <= 0:
        raise ToolError("--timeout must be a positive number of seconds.")
    allowed = drawing_guard(src, args.allow_drawing_rewrite)
    wb = open_workbook(src)
    if args.sheet is not None:
        only = pick_sheet(wb, args.sheet).title
    else:
        only = None
    cells_scope = None
    if args.cells is not None:
        scope_sheet = only if only is not None else pick_sheet(wb, None).title
        cells_scope = _recalc_scope_cells(args.cells, scope_sheet)

    sys.setrecursionlimit(max(sys.getrecursionlimit(), 10000))
    model = build_model(wb)
    formula_keys = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    formula_keys.append((ws.title, cell.row, cell.column))
    candidates, anchors = array_candidates(wb)
    evaluator = Evaluator(model, time.monotonic() + args.timeout, collect_names(wb), candidates, anchors)
    outcomes = {}
    try:
        for key in formula_keys:
            outcomes[key] = evaluator.evaluate_formula(key)
    except Unsupported as exc:
        raise ToolError(_recalc_error_message(exc, "a formula")) from None
    except Circular as exc:
        sheet_name, r, c = exc.key
        raise ToolError(f"Circular reference at {sheet_name}!{coord(r, c)}. Nothing was written.") from None
    except Timeout:
        raise ToolError(f"Recalculation timed out after {fmt_num(args.timeout)} seconds.") from None
    except RecursionError:
        raise ToolError("Formula chain is too deep for the built-in evaluator. Nothing was written.") from None
    results = {key: outcomes[key].value for key in formula_keys}

    def value_of(key):
        if key in outcomes:
            return outcomes[key].value
        return evaluator.cell_value(*key)

    if cells_scope is not None:
        scope = cells_scope
    else:
        scope = [k for k in formula_keys if only is None or k[0] == only]
        # the cells of a spilled block are reported with the formula that owns them
        for key in list(scope):
            block = outcomes[key].block
            if block is not None:
                for i in range(block.height):
                    for j in range(block.width):
                        if i or j:
                            scope.append((key[0], key[1] + i, key[2] + j))
    computed = {}
    errors = {}
    for key in scope:
        value = value_of(key)
        name = f"{key[0]}!{coord(key[1], key[2])}"
        computed[name] = json_value(value)
        if isinstance(value, XLError):
            errors[name] = value.code

    texts = {}
    for key in formula_keys:
        texts[key] = formula_text(wb[key[0]].cell(key[1], key[2]))
    cache_by_sheet = {}
    for key in formula_keys:
        cache_by_sheet.setdefault(key[0], {})[coord(key[1], key[2])] = results[key]
    dynamic_by_sheet = {}
    for key in formula_keys:
        if outcomes[key].array or outcomes[key].block is not None:
            dynamic_by_sheet.setdefault(key[0], set()).add(coord(key[1], key[2]))

    # Rewrite the formulas: old spill cells are emptied, a multi-value result becomes an array
    # formula over its block (anchor) and the block's cached values (other cells).
    clear_array_members(wb)
    for key in formula_keys:
        sheet_name, r, c = key
        ws = wb[sheet_name]
        stripped = texts[key][1:]
        block = outcomes[key].block
        if block is None:
            if outcomes[key].array:
                ws.cell(r, c).value = ArrayFormula(coord(r, c), "=" + stripped)
            else:
                ws.cell(r, c).value = "=" + stripped
            continue
        ref = f"{coord(r, c)}:{coord(r + block.height - 1, c + block.width - 1)}"
        ws.cell(r, c).value = ArrayFormula(ref, "=" + stripped)
        for i in range(block.height):
            for j in range(block.width):
                if i or j:
                    set_spill_value(ws.cell(r + i, c + j), block.rows[i][j])

    buf = io.BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    def save(path):
        inject_cached_values(xlsx_bytes, str(path), cache_by_sheet, dynamic_by_sheet)

    def verify(path):
        wb_v = open_workbook(path, data_only=True)
        wb_f = open_workbook(path)
        for key in formula_keys:
            sheet_name, r, c = key
            ref = coord(r, c)
            expected = results[key]
            got = wb_v[sheet_name].cell(r, c).value
            if isinstance(got, dt.datetime):
                got = to_excel(got)
            if isinstance(expected, XLError):
                ok = got == expected.code
            else:
                ok = values_equal(expected, got) or (expected == 0 and got is None)
            if not ok:
                raise ToolError(f"{sheet_name}!{ref} has cached value {got!r}, expected {expected!r}")
            if formula_text(wb_f[sheet_name].cell(r, c)) != texts[key]:
                raise ToolError(f"{sheet_name}!{ref} formula text changed")
            block = outcomes[key].block
            if block is None:
                continue
            for i in range(block.height):
                for j in range(block.width):
                    if i == 0 and j == 0:
                        continue
                    want = block.rows[i][j]
                    got = wb_v[sheet_name].cell(r + i, c + j).value
                    if isinstance(got, dt.datetime):
                        got = to_excel(got)
                    if isinstance(want, XLError):
                        ok = got == want.code
                    else:
                        ok = values_equal(want, got) or (want == "" and got is None)
                    if not ok:
                        raise ToolError(f"{sheet_name}!{coord(r + i, c + j)} has value {got!r}, "
                                        f"expected {want!r}")

    commit_workbook(save, target, verify)
    payload = {"output": str(target), "sheet": only, "computed": computed, "errors": errors,
               "drawings_rewritten": allowed}
    emit(args, payload, [f"Recalculated {len(formula_keys)} formula cell(s); {len(errors)} error(s). "
                         f"Wrote {target}."])
    return 0


def cmd_errors(args):
    wb_f = open_workbook(args.file)
    wb_v = open_workbook(args.file, data_only=True)
    sheets = [pick_sheet(wb_f, args.sheet)] if args.sheet is not None else wb_f.worksheets
    findings = []
    error_count = 0
    not_calc = 0
    for ws in sheets:
        wv = wb_v[ws.title]
        for row in ws.iter_rows():
            for cell in row:
                ref = coord(cell.row, cell.column)
                if cell.data_type == "f":
                    value = wv.cell(cell.row, cell.column).value
                    if value is None:
                        not_calc += 1
                        findings.append({"sheet": ws.title, "cell": ref,
                                         "formula": formula_text(cell), "value": None})
                    elif isinstance(value, str) and value in ERROR_CODES:
                        error_count += 1
                        findings.append({"sheet": ws.title, "cell": ref,
                                         "formula": formula_text(cell), "value": value})
                elif cell.data_type == "e" and cell.value in ERROR_CODES:
                    error_count += 1
                    findings.append({"sheet": ws.title, "cell": ref, "formula": None,
                                     "value": cell.value})
    payload = {"error_count": error_count, "not_calculated_count": not_calc, "findings": findings}
    lines = [f"{f['sheet']}!{f['cell']}\t{f['value']}\t{f['formula'] or ''}" for f in findings]
    emit(args, payload, lines)
    if error_count:
        warn(f"{error_count} formula error(s) found. See findings.")
    if not_calc:
        warn(f"{not_calc} formula(s) have no calculated value. Run recalc.")
    return 1 if findings else 0


class _CheckFail(Exception):
    pass


def cmd_validate(args):
    path = args.file
    if not Path(path).is_file():
        raise ToolError(f"Input file not found: {path}")
    check_format(path)
    checks = []
    state = {}

    def step_zip():
        try:
            with zipfile.ZipFile(path) as zf:
                bad = zf.testzip()
                state["names"] = zf.namelist()
        except (zipfile.BadZipFile, OSError) as exc:
            raise _CheckFail(f"not a readable zip file (corrupt or truncated): {exc}")
        if bad is not None:
            raise _CheckFail(f"damaged zip member: {bad}")
        return "zip container reads"

    def step_parts():
        names = state["names"]
        missing = [p for p in ("[Content_Types].xml", "xl/workbook.xml") if p not in names]
        if missing:
            raise _CheckFail("missing required part(s): " + ", ".join(missing))
        return "[Content_Types].xml and xl/workbook.xml present"

    def step_xml():
        with zipfile.ZipFile(path) as zf:
            for name in state["names"]:
                if name.endswith((".xml", ".rels")):
                    try:
                        etree.fromstring(zf.read(name), SAFE_PARSER)
                    except etree.XMLSyntaxError as exc:
                        raise _CheckFail(f"{name} does not parse: {exc}")
        return "all XML and relationship parts parse"

    def step_rels():
        names = set(state["names"])
        with zipfile.ZipFile(path) as zf:
            for name in state["names"]:
                if not name.endswith(".rels"):
                    continue
                root = etree.fromstring(zf.read(name), SAFE_PARSER)
                folder = posixpath.dirname(name)
                base = posixpath.dirname(folder)
                for rel in root:
                    if rel.get("TargetMode") == "External":
                        continue
                    target = rel.get("Target") or ""
                    if target.startswith("/"):
                        resolved = target[1:]
                    else:
                        resolved = posixpath.normpath(posixpath.join(base, target))
                    if resolved.startswith("..") or resolved not in names:
                        raise _CheckFail(f"{name} points to a missing part: {target}")
        return "all internal relationship targets exist"

    def step_openpyxl():
        try:
            wb = openpyxl.load_workbook(path)
        except Exception as exc:
            raise _CheckFail(f"openpyxl cannot open the workbook: {type(exc).__name__}")
        state["sheetnames"] = list(wb.sheetnames)
        return "openpyxl opens the workbook"

    def step_has_sheet():
        if not state["sheetnames"]:
            raise _CheckFail("the workbook has no sheets")
        return f"{len(state['sheetnames'])} sheet(s)"

    def step_sheet_names():
        names = state["sheetnames"]
        for name in names:
            if len(name) > 31:
                raise _CheckFail(f'sheet name "{name}" is longer than 31 characters')
            bad = [ch for ch in name if ch in "\\/?*[]:"]
            if bad:
                raise _CheckFail(f'sheet name "{name}" contains {"".join(bad)}')
        folded = [n.casefold() for n in names]
        if len(set(folded)) != len(folded):
            raise _CheckFail("sheet names are not unique")
        return "sheet names are unique, at most 31 characters, no forbidden characters"

    steps = [("readable zip", step_zip), ("required parts", step_parts),
             ("XML parts parse", step_xml), ("relationship targets", step_rels),
             ("openpyxl opens workbook", step_openpyxl), ("at least one sheet", step_has_sheet),
             ("sheet names", step_sheet_names)]
    failed = False
    for name, fn in steps:
        if failed:
            checks.append({"check": name, "result": "not run", "detail": ""})
            continue
        try:
            detail = fn()
            checks.append({"check": name, "result": "pass", "detail": detail})
        except _CheckFail as exc:
            checks.append({"check": name, "result": "fail", "detail": str(exc)})
            failed = True
    verdict = "invalid" if failed else "valid"
    payload = {"file": path, "checks": checks, "verdict": verdict}
    lines = [f"{c['result'].upper()}\t{c['check']}\t{c['detail']}" for c in checks]
    lines.append(f"Verdict: {verdict}")
    emit(args, payload, lines)
    if failed:
        warn("The workbook failed validation.")
        return 1
    return 0


def load_rules(path):
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        raise ToolError(f"Rules file is invalid: file not found: {path}.") from None
    except (OSError, ValueError) as exc:
        raise ToolError(f"Rules file is invalid: {exc}.") from None

    def bad(reason):
        raise ToolError(f"Rules file is invalid: {reason}.")

    if not isinstance(data, dict):
        bad('the top level must be an object with a "columns" key')
    for key in data:
        if key != "columns":
            bad(f'unknown key "{key}"')
    columns = data.get("columns")
    if not isinstance(columns, dict):
        bad('"columns" must be an object that maps header text to rules')
    rules = []
    for name, rule in columns.items():
        if not isinstance(rule, dict):
            bad(f'the rule for column "{name}" must be an object')
        for key in rule:
            if key not in CHECK_RULE_KEYS:
                bad(f'unknown rule key "{key}" in column "{name}"')
        if rule.get("type") not in CHECK_TYPES:
            bad(f'column "{name}" needs a "type" of integer, decimal, date, boolean or text')
        kind = rule["type"]
        compiled = dict(name=name, type=kind, required=False, unique=False)
        for flag in ("required", "unique"):
            if flag in rule:
                if not isinstance(rule[flag], bool):
                    bad(f'"{flag}" in column "{name}" must be true or false')
                compiled[flag] = rule[flag]
        for bound in ("min", "max"):
            if bound in rule:
                if kind not in ("integer", "decimal", "date"):
                    bad(f'"{bound}" in column "{name}" applies only to integer, decimal or date')
                raw = rule[bound]
                if kind == "date":
                    if not (isinstance(raw, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw)):
                        bad(f'"{bound}" in column "{name}" must be an ISO YYYY-MM-DD date')
                    try:
                        compiled[bound] = dt.datetime.strptime(raw, "%Y-%m-%d")
                    except ValueError:
                        bad(f'"{bound}" in column "{name}" is not a real date')
                else:
                    if not is_number(raw):
                        bad(f'"{bound}" in column "{name}" must be a number')
                    compiled[bound] = raw
        if "allowed" in rule:
            if not isinstance(rule["allowed"], list):
                bad(f'"allowed" in column "{name}" must be a list')
            compiled["allowed"] = rule["allowed"]
        if "pattern" in rule:
            if not isinstance(rule["pattern"], str):
                bad(f'"pattern" in column "{name}" must be text')
            try:
                compiled["pattern"] = re.compile(rule["pattern"])
            except re.error as exc:
                bad(f'"pattern" in column "{name}" is not a valid regular expression ({exc})')
        if "max_length" in rule:
            ml = rule["max_length"]
            if not isinstance(ml, int) or isinstance(ml, bool) or ml < 0:
                bad(f'"max_length" in column "{name}" must be a whole number of 0 or more')
            compiled["max_length"] = ml
        rules.append(compiled)
    return rules


def _is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _type_ok(v, kind):
    if kind == "integer":
        return is_number(v) and float(v).is_integer()
    if kind == "decimal":
        return is_number(v)
    if kind == "date":
        return isinstance(v, (dt.date, dt.datetime))
    if kind == "boolean":
        return isinstance(v, bool)
    return isinstance(v, str)


def _text_form(v):
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if is_number(v):
        return num_to_text(v)
    return json_value(v) if isinstance(v, (dt.date, dt.datetime)) else str(v)


def _as_datetime(v):
    if isinstance(v, dt.datetime):
        return v
    return dt.datetime(v.year, v.month, v.day)


def _norm_compare(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return _text_form(v)
    return v


def _allowed_match(value, allowed):
    v = _norm_compare(value)
    a = _norm_compare(allowed)
    if isinstance(v, bool) or isinstance(a, bool):
        return isinstance(v, bool) and isinstance(a, bool) and v == a
    if is_number(v) and is_number(a):
        return v == a
    if isinstance(v, str) and isinstance(a, str):
        return v == a
    return False


def _unique_key(v):
    if isinstance(v, bool):
        return ("b", v)
    if is_number(v):
        return ("n", float(v))
    if isinstance(v, (dt.date, dt.datetime)):
        return ("d", _text_form(v))
    return ("s", v)


def check_cell(value, rule, seen):
    """Return the names of the rules this value breaks (empty list when it passes)."""
    if _is_empty(value):
        return ["required"] if rule["required"] else []
    kind = rule["type"]
    if not _type_ok(value, kind):
        return ["type"]
    fails = []
    if "min" in rule:
        cmp_value = _as_datetime(value) if kind == "date" else value
        if cmp_value < rule["min"]:
            fails.append("min")
    if "max" in rule:
        cmp_value = _as_datetime(value) if kind == "date" else value
        if cmp_value > rule["max"]:
            fails.append("max")
    if "allowed" in rule and not any(_allowed_match(value, a) for a in rule["allowed"]):
        fails.append("allowed")
    if "pattern" in rule and not rule["pattern"].fullmatch(_text_form(value)):
        fails.append("pattern")
    if "max_length" in rule and len(_text_form(value)) > rule["max_length"]:
        fails.append("max_length")
    if rule["unique"]:
        key = _unique_key(value)
        if key in seen:
            fails.append("unique")
        seen.add(key)
    return fails


def cmd_check_values(args):
    if args.header_row < 1:
        raise ToolError("--header-row must be 1 or more.")
    rules = load_rules(args.rules)
    wb_f = open_workbook(args.file)
    wb_v = open_workbook(args.file, data_only=True)
    ws_f = pick_sheet(wb_f, args.sheet)
    ws_v = wb_v[ws_f.title]
    headers = {}
    for c in range(1, ws_v.max_column + 1):
        value = ws_v.cell(args.header_row, c).value
        if value is None:
            continue
        headers.setdefault(str(value), c)
    for rule in rules:
        if rule["name"] not in headers:
            raise ToolError(f'Column "{rule["name"]}" named in rules was not found in row {args.header_row}.')
    seen = {rule["name"]: set() for rule in rules}
    violations = []
    rows_checked = 0
    for r in range(args.header_row + 1, ws_v.max_row + 1):
        values = {rule["name"]: ws_v.cell(r, headers[rule["name"]]).value for rule in rules}
        if all(_is_empty(v) for v in values.values()):
            continue
        rows_checked += 1
        for rule in rules:
            name = rule["name"]
            col = headers[name]
            value = values[name]
            ref = coord(r, col)
            if value is None and ws_f.cell(r, col).data_type == "f":
                violations.append({"sheet": ws_v.title, "cell": ref, "column": name,
                                   "rule": "not_calculated", "value": None})
                continue
            for failed in check_cell(value, rule, seen[name]):
                violations.append({"sheet": ws_v.title, "cell": ref, "column": name,
                                   "rule": failed, "value": json_value(value)})
    payload = {"sheet": ws_v.title, "rows_checked": rows_checked, "violations": violations}
    lines = [f"{v['cell']}\t{v['column']}\t{v['rule']}\t{v['value']}" for v in violations]
    emit(args, payload, lines)
    if violations:
        warn(f"{len(violations)} value(s) failed the rules. See violations.")
        return 1
    return 0


# ---------------------------------------------------------------- argument parsing

def build_parser():
    parser = _Parser(prog="xlsx_tool.py",
                     description="Read, write, format, recalculate, check and validate .xlsx workbooks.")
    sub = parser.add_subparsers(dest="command", required=True, parser_class=_Parser, metavar="COMMAND")

    p = sub.add_parser("sheets", help="list sheets")
    p.add_argument("file")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_sheets)

    p = sub.add_parser("read", help="read cells or ranges")
    p.add_argument("file")
    p.add_argument("--sheet")
    p.add_argument("--range")
    p.add_argument("--mode", choices=["values", "formulas"], default="values")
    p.add_argument("--out")
    p.add_argument("--force", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_read)

    p = sub.add_parser("write", help="write values and formulas from a patch file")
    p.add_argument("file")
    p.add_argument("--patch", required=True)
    p.add_argument("--out")
    p.add_argument("--force", action="store_true")
    p.add_argument("--in-place", action="store_true")
    p.add_argument("--allow-drawing-rewrite", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_write)

    p = sub.add_parser("format", help="apply a number format to a range")
    p.add_argument("file")
    p.add_argument("--sheet")
    p.add_argument("--range")
    p.add_argument("--preset", choices=["currency", "date", "percent"])
    p.add_argument("--code")
    p.add_argument("--currency-symbol", default="$")
    p.add_argument("--out")
    p.add_argument("--force", action="store_true")
    p.add_argument("--in-place", action="store_true")
    p.add_argument("--allow-drawing-rewrite", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_format)

    p = sub.add_parser("recalc", help="recalculate formulas and store the results")
    p.add_argument("file")
    p.add_argument("--sheet")
    p.add_argument("--cells")
    p.add_argument("--out")
    p.add_argument("--force", action="store_true")
    p.add_argument("--in-place", action="store_true")
    p.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT)
    p.add_argument("--allow-drawing-rewrite", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_recalc)

    p = sub.add_parser("errors", help="report formula errors and uncalculated formulas")
    p.add_argument("file")
    p.add_argument("--sheet")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_errors)

    p = sub.add_parser("validate", help="check the file is a valid workbook")
    p.add_argument("file")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_validate)

    p = sub.add_parser("check-values", help="check test data against a rules file")
    p.add_argument("file")
    p.add_argument("--rules", required=True)
    p.add_argument("--sheet")
    p.add_argument("--header-row", type=int, default=1)
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_check_values)
    return parser


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except ToolError as exc:
        warn(str(exc))
        return exc.exit_code
    except PermissionError:
        warn("The file is open in another program. Close it and try again.")
        return 1
    except Exception as exc:  # any unexpected failure: nothing is kept
        warn(f"Internal error ({type(exc).__name__}). No output was kept.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
