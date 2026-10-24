#!/usr/bin/env python3
"""docx_tool.py - one command-line entry point for reading and changing .docx files.

Usage: python docx_tool.py <command> <input> [options]
Run `python docx_tool.py --help` for the command list. See SKILL.md for every command.
Exit codes: 0 done, 1 error (bad input or option, nothing written), 2 out of scope.
"""

from __future__ import annotations

import argparse
import copy
import importlib.util
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % WML
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"

CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

NOT_SEARCHED = ["footnotes", "endnotes", "comments", "field codes"]

LEGACY_MSG = "Legacy .doc is not supported. Save the file as .docx and try again."
PASSWORD_MSG = "This Word file is password-protected. The skill does not open protected files."
MACRO_MSG = "Macro-enabled .docm files are not supported. Save as .docx (macros are dropped)."
UNSUPPORTED_MSG = "Formatting or paragraph-mark tracked changes cannot be resolved by this skill."
LIBREOFFICE_MSG = "LibreOffice is not installed. Install it or convert the file in Word."
LOCKED_MSG = "The file is open in another program. Close it and try again."
REFUSE_INPUT_MSG = "Refusing to overwrite the input. Use --in-place to change it."
CORRUPT_MSG = "File is not a readable .docx (corrupt or truncated): {}"

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ENCRYPTION_NAME = "EncryptionInfo".encode("utf-16-le")
_PLACEHOLDER = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")
_HEADING = re.compile(r"Heading ([1-9])")
_NUMBERED = re.compile(r"\d+\.\s+(.*)")

# Elements whose content is never visible body text for our purposes.
_NO_TEXT_TAGS = {
    W + "pPr", W + "rPr", W + "del", W + "moveFrom", W + "delText", W + "instrText",
    W + "delInstrText", W + "drawing", W + "pict", W + "object", W + "txbxContent",
}
_TEXT_TAGS = {W + "tab": "\t", W + "br": "\n", W + "cr": "\n", W + "noBreakHyphen": "-"}

_CHANGE_TAGS = {
    "rPrChange": "formatting change", "pPrChange": "formatting change",
    "sectPrChange": "formatting change", "tblPrChange": "formatting change",
    "trPrChange": "formatting change", "tcPrChange": "formatting change",
    "tblGridChange": "formatting change", "numberingChange": "formatting change",
    "cellIns": "table cell revision", "cellDel": "table cell revision",
    "moveFrom": "move revision", "moveTo": "move revision",
}
_CONTENT_PARENTS = {"p", "hyperlink", "smartTag", "sdtContent", "customXml", "fldSimple", "dir", "bdo"}


class ToolError(Exception):
    """A message for stderr and an exit code. Nothing is written when raised."""

    def __init__(self, code: int, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


class VerifyError(Exception):
    """The re-opened output failed its check."""


class DocxArgumentParser(argparse.ArgumentParser):
    def error(self, message):  # argparse usage errors exit 1, not 2
        raise ToolError(1, f"{self.format_usage().strip()}\n{message}")


# ---------------------------------------------------------------- input checks

def precheck(path: Path) -> None:
    """Refuse missing and out-of-scope inputs before any parsing."""
    if not path.is_file():
        raise ToolError(1, f"Input file not found: {path}")
    suffix = path.suffix.lower()
    if suffix == ".doc":
        raise ToolError(2, LEGACY_MSG)
    if suffix == ".docm":
        raise ToolError(2, MACRO_MSG)
    with path.open("rb") as fh:
        head = fh.read(8)
    if head == _OLE_MAGIC:
        data = path.read_bytes()
        raise ToolError(2, PASSWORD_MSG if _ENCRYPTION_NAME in data else LEGACY_MSG)
    if zipfile.is_zipfile(path):
        try:
            with zipfile.ZipFile(path) as zf:
                content_types = zf.read("[Content_Types].xml") if "[Content_Types].xml" in zf.namelist() else b""
        except PermissionError:
            raise
        except Exception:  # unreadable zip is reported later as corrupt
            content_types = b""
        if b"macroEnabled" in content_types:
            raise ToolError(2, MACRO_MSG)


def open_docx(path: Path) -> Document:
    """Return an opened python-docx Document or raise the matching ToolError."""
    precheck(path)
    try:
        with zipfile.ZipFile(path) as zf:
            bad_member = zf.testzip()
    except PermissionError:
        raise
    except Exception:
        bad_member = "unreadable"
    if bad_member is not None:
        raise ToolError(1, CORRUPT_MSG.format(path))
    try:
        return Document(str(path))
    except PermissionError:
        raise
    except Exception:
        raise ToolError(1, CORRUPT_MSG.format(path))


def reopen(path: Path) -> Document:
    try:
        return Document(str(path))
    except Exception as exc:
        raise VerifyError(f"{type(exc).__name__}: {exc}") from None


# ----------------------------------------------------------- output handling

def same_file(a: Path, b: Path) -> bool:
    if os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b)):
        return True
    try:
        return a.exists() and os.path.samefile(a, b)
    except OSError:
        return False


def resolve_target(args, input_path: Path, command: str, ext: str = ".docx") -> tuple[Path, bool]:
    """Return (target path, in_place flag) and refuse bad targets before any work."""
    in_place = bool(getattr(args, "in_place", False))
    out = getattr(args, "out", None)
    if in_place and out:
        raise ToolError(1, "Use either --out or --in-place, not both.")
    if in_place:
        return input_path, True
    target = Path(out) if out else input_path.with_name(f"{input_path.stem}.{command}{ext}")
    if same_file(target, input_path):
        raise ToolError(1, REFUSE_INPUT_MSG)
    if target.exists() and not getattr(args, "force", False):
        raise ToolError(1, f"Output already exists: {target}. Choose another --out or pass --force.")
    if not target.parent.is_dir():
        raise ToolError(1, f"Output folder does not exist: {target.parent}")
    return target, False


def finalize(doc: Document, target: Path, verify):
    """Save to a temp file in the target folder, re-open it, then move it into place."""
    fd, tmp_name = tempfile.mkstemp(prefix=".docx-tool-", suffix=".docx.tmp", dir=str(target.parent))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        doc.save(str(tmp))
        try:
            info = verify(tmp)
        except VerifyError as exc:
            raise ToolError(1, f"Output failed verification and was not kept: {exc}") from None
        os.replace(str(tmp), str(target))
        return info
    finally:
        if tmp.exists():
            tmp.unlink()


def emit(args, data: dict, lines: list[str]) -> None:
    if getattr(args, "json", False):
        sys.stdout.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
    else:
        sys.stdout.write("\n".join(lines) + "\n")


def report_write(args, data: dict, lines: list[str], searched: bool = False) -> None:
    """Print the result of a write command. `searched` adds the not-searched list."""
    if searched:
        data["not_searched"] = NOT_SEARCHED
        lines = lines + [f"Not searched: {', '.join(NOT_SEARCHED)}"]
    emit(args, data, lines)


# ------------------------------------------------------- paragraph text model

@dataclass
class Piece:
    """One visible unit of text: a w:t, or a tab/break, with the run that holds it."""
    el: etree._Element
    run: etree._Element
    text: str
    tracked: bool
    start: int = 0

    @property
    def end(self) -> int:
        return self.start + len(self.text)

    @property
    def is_text(self) -> bool:
        return self.el.tag == W + "t"


def _collect(el, tracked: bool, run, out: list) -> None:
    for child in el:
        tag = child.tag
        if not isinstance(tag, str) or tag in _NO_TEXT_TAGS:
            continue
        if tag == W + "r":
            _collect(child, tracked, child, out)
        elif tag in (W + "ins", W + "moveTo"):
            _collect(child, True, run, out)
        elif tag == W + "t" or tag in _TEXT_TAGS:
            if run is None:
                continue
            text = (child.text or "") if tag == W + "t" else _TEXT_TAGS[tag]
            if text:
                out.append(Piece(child, run, text, tracked))
        else:
            _collect(child, tracked, run, out)


def build_model(p) -> list[Piece]:
    pieces: list[Piece] = []
    _collect(p, False, None, pieces)
    pos = 0
    for piece in pieces:
        piece.start = pos
        pos += len(piece.text)
    return pieces


def model_text(model: list[Piece]) -> str:
    return "".join(piece.text for piece in model)


def para_text(p) -> str:
    return model_text(build_model(p))


def _overlaps_tracked(model: list[Piece], start: int, end: int) -> bool:
    return any(pc.tracked and pc.start < end and pc.end > start for pc in model)


def rewrite_span(model: list[Piece], start: int, end: int, new: str) -> None:
    """Put `new` into the first run the span touches; remove the rest of the span.

    `model` must be fresh (built from the current XML).
    """
    touched = [pc for pc in model if pc.start < end and pc.end > start]
    for index, pc in enumerate(touched):
        a = max(start, pc.start) - pc.start
        b = min(end, pc.end) - pc.start
        if pc.is_text:
            pc.el.text = pc.text[:a] + (new if index == 0 else "") + pc.text[b:]
            pc.el.set(XML_SPACE, "preserve")
        elif index == 0 and new:
            fresh = OxmlElement("w:t")
            fresh.text = new
            fresh.set(XML_SPACE, "preserve")
            pc.el.addprevious(fresh)
            pc.el.getparent().remove(pc.el)
        else:
            pc.el.getparent().remove(pc.el)


def find_all(text: str, needle: str) -> list[int]:
    found, i = [], 0
    while True:
        j = text.find(needle, i)
        if j < 0:
            return found
        found.append(j)
        i = j + len(needle)


# ------------------------------------------------------------- document parts

def story_roots(doc: Document) -> list[tuple[str, str, etree._Element]]:
    """(location, part name, root element) for the body and every header and footer."""
    roots = [("body", "", doc.element.body)]
    extra = []
    for part in doc.part.package.iter_parts():
        if part.content_type == CT_HEADER:
            extra.append(("header", str(part.partname), part.element))
        elif part.content_type == CT_FOOTER:
            extra.append(("footer", str(part.partname), part.element))
    extra.sort(key=lambda item: item[1])
    return roots + extra


def comments_root(doc: Document):
    for part in doc.part.package.iter_parts():
        if part.content_type == CT_COMMENTS:
            return part.element
    return None


def block_items(container):
    """Yield ('p', el) or ('tbl', el) in order, looking inside block-level content controls."""
    for child in container:
        if child.tag == W + "p":
            yield "p", child
        elif child.tag == W + "tbl":
            yield "tbl", child
        elif child.tag == W + "sdt":
            content = child.find(W + "sdtContent")
            if content is not None:
                yield from block_items(content)


def table_rows(tbl) -> list[list[str]]:
    rows = []
    for tr in tbl.findall(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            cells.append("\n".join(para_text(p) for kind, p in block_items(tc) if kind == "p"))
        rows.append(cells)
    return rows


def table_size(tbl) -> list[int]:
    rows = tbl.findall(W + "tr")
    grid = tbl.find(W + "tblGrid")
    if grid is not None and len(grid):
        cols = len(grid.findall(W + "gridCol"))
    else:
        cols = max((len(r.findall(W + "tc")) for r in rows), default=0)
    return [len(rows), cols]


def container_text(container) -> str:
    lines = []
    for kind, el in block_items(container):
        if kind == "p":
            lines.append(para_text(el))
        else:
            lines.extend("\t".join(row) for row in table_rows(el))
    return "\n".join(lines)


def style_name(p, doc: Document) -> str:
    style = Paragraph(p, doc).style
    return (style.name if style is not None and style.name else "Normal")


def count_paragraphs(roots) -> int:
    return sum(len(list(root.iter(W + "p"))) for _loc, _name, root in roots)


# ---------------------------------------------------------------- revisions

@dataclass
class Revision:
    el: etree._Element
    kind: str            # insert | delete | unsupported
    rid: int | None
    author: str
    date: str | None
    text: str
    location: str
    paragraph_index: int | None
    detail: str | None


def _int_or_none(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _classify(el, local: str):
    if local in _CHANGE_TAGS:
        return "unsupported", _CHANGE_TAGS[local]
    if any(etree.QName(a).localname in ("ins", "del") for a in el.iterancestors()):
        return "unsupported", "nested inside another revision"
    parent = el.getparent()
    parent_local = etree.QName(parent).localname if parent is not None else ""
    if parent_local in _CONTENT_PARENTS:
        return ("insert" if local == "ins" else "delete"), None
    if parent_local == "rPr" and parent.getparent() is not None \
            and etree.QName(parent.getparent()).localname == "pPr":
        return "unsupported", "paragraph mark revision"
    return "unsupported", f"revision mark inside w:{parent_local}"


def _revision_text(el, kind: str) -> str:
    if kind == "insert":
        return "".join(t.text or "" for t in el.iter(W + "t"))
    if kind == "delete":
        return "".join(t.text or "" for t in el.iter(W + "delText", W + "t"))
    return ""


def scan_revisions(roots) -> list[Revision]:
    found: list[Revision] = []
    tags = [W + name for name in ("ins", "del", *_CHANGE_TAGS)]
    for location, _name, root in roots:
        index = {p: i for i, p in enumerate(root.iter(W + "p"))}
        for el in root.iter(*tags):
            local = etree.QName(el).localname
            kind, detail = _classify(el, local)
            rid = _int_or_none(el.get(W + "id"))
            if kind != "unsupported" and rid is None:
                kind, detail = "unsupported", "revision has no valid w:id"
            para = next(el.iterancestors(W + "p"), None)
            found.append(Revision(
                el, kind, rid, el.get(W + "author") or "", el.get(W + "date"),
                _revision_text(el, kind), location,
                index.get(para) if para is not None else None, detail,
            ))
    return found


def _unwrap(el) -> None:
    parent = el.getparent()
    position = parent.index(el)
    for child in list(el):
        parent.insert(position, child)
        position += 1
    parent.remove(el)


def _comment_marks_inside(el) -> dict:
    """Comment id -> {'start', 'end', 'ref'} for the comment markers that lie inside `el`."""
    marks: dict[str, dict] = {}
    for node in el.iter(W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference"):
        slot = marks.setdefault(node.get(W + "id") or "", {"start": None, "end": None, "ref": None})
        if node.tag == W + "commentRangeStart":
            slot["start"] = node
        elif node.tag == W + "commentRangeEnd":
            slot["end"] = node
        else:
            slot["ref"] = node.getparent()
    return marks


def _release_comment_marks(el) -> None:
    """Before a revision element is removed, move the comment markers inside it out of it.

    A start marker goes before the element; an end marker goes after it, followed by its
    reference run. The comment's text is removed with the element, so the markers are kept
    until `_drop_comment` decides whether the comment itself goes.
    """
    trail = el
    for _cid, marks in _comment_marks_inside(el).items():
        start, end, ref = marks["start"], marks["end"], marks["ref"]
        if start is not None:
            el.addprevious(start)
        if ref is not None and end is None:
            el.addprevious(ref)
        if end is not None:
            trail.addnext(end)
            trail = end
            if ref is not None:
                trail.addnext(ref)
                trail = ref


_CONTENT_TAGS = {W + "t", W + "delText", W + "tab", W + "br", W + "cr", W + "noBreakHyphen"}


def _comments_with_content(body) -> set[str]:
    """Ids of comments whose range (start marker to end marker) holds text, deleted text included."""
    open_ranges: dict[str, bool] = {}
    with_content: set[str] = set()
    for el in body.iter():
        tag = el.tag
        if not isinstance(tag, str):
            continue
        if tag == W + "commentRangeStart":
            open_ranges[el.get(W + "id") or ""] = False
        elif tag == W + "commentRangeEnd":
            cid = el.get(W + "id") or ""
            if open_ranges.pop(cid, False):
                with_content.add(cid)
        elif tag in _CONTENT_TAGS:
            for cid in open_ranges:
                open_ranges[cid] = True
    return with_content


def _drop_comment(body, comments, cid: str) -> None:
    """Remove a comment: its range markers, its reference run and its entry in the comments part."""
    for el in list(body.iter(W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference")):
        if (el.get(W + "id") or "") != cid:
            continue
        target = el.getparent() if el.tag == W + "commentReference" else el
        if target.getparent() is not None:
            target.getparent().remove(target)
    if comments is not None:
        for c in comments.findall(W + "comment"):
            if (c.get(W + "id") or "") == cid:
                comments.remove(c)


def apply_revision(el, kind: str, mode: str) -> None:
    keep = (kind == "insert" and mode == "accept") or (kind == "delete" and mode == "reject")
    if not keep:
        _release_comment_marks(el)
        el.getparent().remove(el)
        return
    if kind == "delete":  # reject a deletion: the deleted text becomes normal text again
        for node in list(el.iter(W + "delText")):
            node.tag = W + "t"
    _unwrap(el)


# ------------------------------------------------------------ comment helpers

def _in_deleted(el) -> bool:
    return any(a.tag in (W + "del", W + "moveFrom") for a in el.iterancestors())


def anchor_texts(body) -> dict[str, str]:
    """Map comment id to the text between its commentRangeStart and commentRangeEnd."""
    open_ranges: dict[str, list[str]] = {}
    result: dict[str, str] = {}
    for el in body.iter():
        tag = el.tag
        if not isinstance(tag, str):
            continue
        if tag == W + "commentRangeStart":
            open_ranges[el.get(W + "id")] = []
        elif tag == W + "commentRangeEnd":
            cid = el.get(W + "id")
            result[cid] = "".join(open_ranges.pop(cid, []))
        elif tag == W + "p":
            for parts in open_ranges.values():
                parts.append("\n")
        elif open_ranges and (tag == W + "t" or tag in _TEXT_TAGS) and not _in_deleted(el):
            if tag == W + "t":
                piece = el.text or ""
            else:
                piece = _TEXT_TAGS[tag]
            for parts in open_ranges.values():
                parts.append(piece)
    return result


def list_comments(doc: Document) -> list[dict]:
    root = comments_root(doc)
    if root is None:
        return []
    anchors = anchor_texts(doc.element.body)
    rows = []
    for c in root.findall(W + "comment"):
        cid = c.get(W + "id") or ""
        rows.append({
            "id": int(cid) if cid.isdigit() else cid,
            "author": c.get(W + "author") or "",
            "initials": c.get(W + "initials") or "",
            "date": c.get(W + "date"),
            "anchor_text": anchors.get(cid, ""),
            "text": "\n".join(para_text(p) for p in c.findall(W + "p")),
        })
    return rows


def _split_text(t, k: int):
    text = t.text or ""
    tail = OxmlElement("w:t")
    tail.text = text[k:]
    tail.set(XML_SPACE, "preserve")
    t.text = text[:k]
    t.set(XML_SPACE, "preserve")
    t.addnext(tail)
    return tail


def _split_run_before(run, child) -> None:
    if child.getparent() is not run:
        raise RuntimeError("run layout not supported for comment anchoring")
    new_run = OxmlElement("w:r")
    rpr = run.find(W + "rPr")
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    children = list(run)
    for node in children[children.index(child):]:
        new_run.append(node)
    run.addnext(new_run)


def _ensure_boundary(p, offset: int) -> None:
    """Split runs so that a run boundary sits at `offset` in the paragraph text."""
    for pc in build_model(p):
        if pc.start < offset < pc.end and pc.is_text:
            tail = _split_text(pc.el, offset - pc.start)
            _split_run_before(pc.run, tail)
            return
        if pc.start == offset:
            first_child = next((c for c in pc.run if c.tag != W + "rPr"), None)
            if first_child is not pc.el:
                _split_run_before(pc.run, pc.el)
            return


def _anchored_runs(p, start: int, end: int) -> list:
    runs: list = []
    for pc in build_model(p):
        if pc.start >= start and pc.end <= end and not any(pc.run is r for r in runs):
            runs.append(pc.run)
    return runs


_REVISION_WRAPPERS = {W + "ins", W + "del", W + "moveFrom", W + "moveTo"}


def _lift_out_of_revisions(el, before: bool) -> None:
    """Move `el` out of every tracked-change wrapper around it, before or after the outermost one."""
    wrapper = next((a for a in el.iterancestors() if a.tag in _REVISION_WRAPPERS), None)
    while wrapper is not None:
        if before:
            wrapper.addprevious(el)
        else:
            wrapper.addnext(el)
        wrapper = next((a for a in el.iterancestors() if a.tag in _REVISION_WRAPPERS), None)


def _keep_markers_outside_revisions(body, comment_id) -> None:
    """Move a new comment's range markers and reference run out of any tracked-change wrapper.

    Accepting or rejecting a revision unwraps or removes its w:ins/w:del element with everything
    inside it. A marker left inside would be lost, and Word then shows no comment at all (seen
    after a reject-all). The start marker goes before its wrapper. The end marker goes after it,
    and the reference run goes directly after the end marker, which is the layout Word writes
    itself (Word collapses the comment range when the reference sits before the end marker).
    """
    cid = str(comment_id)
    start = end = ref = None
    for el in body.iter(W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference"):
        if el.get(W + "id") != cid:
            continue
        if el.tag == W + "commentRangeStart":
            start = el
        elif el.tag == W + "commentRangeEnd":
            end = el
        else:
            ref = el.getparent()
    if start is None or end is None:
        return
    _lift_out_of_revisions(start, before=True)
    _lift_out_of_revisions(end, before=False)
    if ref is not None:
        end.addnext(ref)


# ------------------------------------------------------------- replace / fill

class Tally:
    def __init__(self):
        self.done = 0
        self.skipped = 0


def _replace_in_paragraph(p, needle: str, new: str, tally: Tally) -> None:
    for start in reversed(find_all(para_text(p), needle)):
        end = start + len(needle)
        model = build_model(p)
        if model_text(model)[start:end] != needle:
            raise RuntimeError("paragraph text changed during replace")
        if _overlaps_tracked(model, start, end):
            tally.skipped += 1
            continue
        rewrite_span(model, start, end, new)
        tally.done += 1


def replace_everywhere(roots, needle: str, new: str) -> Tally:
    tally = Tally()
    for _loc, _name, root in roots:
        for p in list(root.iter(W + "p")):
            _replace_in_paragraph(p, needle, new, tally)
    return tally


class FillState:
    def __init__(self):
        self.filled = 0
        self.skipped = 0
        self.per_name: dict[str, int] = {}
        self.unfilled: list[str] = []
        self.left_behind = 0

    def miss(self, name: str) -> None:
        self.left_behind += 1
        if name not in self.unfilled:
            self.unfilled.append(name)


def _fill_paragraph(p, values: dict[str, str], state: FillState) -> None:
    pos = 0
    while True:
        model = build_model(p)
        match = _PLACEHOLDER.search(model_text(model), pos)
        if match is None:
            return
        name, start, end = match.group(1), match.start(), match.end()
        if name not in values:
            state.miss(name)
            pos = end
            continue
        if _overlaps_tracked(model, start, end):
            state.skipped += 1
            state.miss(name)
            pos = end
            continue
        value = values[name]
        rewrite_span(model, start, end, value)
        state.per_name[name] = state.per_name.get(name, 0) + 1
        state.filled += 1
        pos = start + len(value)


def count_placeholders(roots) -> int:
    return sum(len(_PLACEHOLDER.findall(para_text(p)))
               for _loc, _name, root in roots for p in root.iter(W + "p"))


def _read_json_object(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ToolError(1, f"Could not read JSON in {path}: {exc}") from None
    if not isinstance(data, dict):
        raise ToolError(1, f"Could not read JSON in {path}: expected a JSON object.")
    return data


def _fill_values(path: Path) -> dict[str, str]:
    values = {}
    for name, value in _read_json_object(path).items():
        if isinstance(value, (dict, list)) or value is None:
            raise ToolError(1, f"Value for {name} must be a text, number or boolean.")
        text = str(value)
        if "{{" in text:
            raise ToolError(1, f'Value for {name} must not contain "{{{{".')
        values[name] = text
    return values


# ----------------------------------------------------------------- commands

def cmd_read(args) -> None:
    doc = open_docx(Path(args.input))
    roots = story_roots(doc)
    paragraphs, tables, order, lines = [], [], [], []
    for kind, el in block_items(doc.element.body):
        if kind == "p":
            entry = {"style": style_name(el, doc), "text": para_text(el)}
            paragraphs.append(entry)
            order.append({"kind": "paragraph", "index": len(paragraphs) - 1})
            lines.append(entry["text"])
        elif not args.no_tables:
            rows = table_rows(el)
            tables.append(rows)
            order.append({"kind": "table", "index": len(tables) - 1})
            lines.extend("\t".join(cell.replace("\n", " ") for cell in row) for row in rows)
            lines.append("")
    data = {
        "paragraphs": paragraphs,
        "tables": tables,
        "order": order,
        "headers": [container_text(root) for loc, _n, root in roots if loc == "header"],
        "footers": [container_text(root) for loc, _n, root in roots if loc == "footer"],
    }
    emit(args, data, lines)


def cmd_summary(args) -> None:
    doc = open_docx(Path(args.input))
    headings, styles, sizes = [], {}, []
    paragraph_count = 0
    for kind, el in block_items(doc.element.body):
        if kind == "p":
            paragraph_count += 1
            name = style_name(el, doc)
            styles[name] = styles.get(name, 0) + 1
            match = _HEADING.fullmatch(name)
            if match:
                headings.append({"level": int(match.group(1)), "text": para_text(el)})
        else:
            sizes.append(table_size(el))
    roots = story_roots(doc)
    root = comments_root(doc)
    has_comments = root is not None and len(root.findall(W + "comment")) > 0
    has_tracked = bool(scan_revisions(roots))
    data = {
        "headings": headings,
        "paragraph_count": paragraph_count,
        "table_count": len(sizes),
        "table_sizes": sizes,
        "styles_used": styles,
        "section_count": len(doc.sections),
        "has_comments": has_comments,
        "has_tracked_changes": has_tracked,
    }
    lines = [
        f"Sections: {data['section_count']}",
        f"Paragraphs: {paragraph_count}",
        f"Tables: {len(sizes)}",
    ]
    lines += [f"  table {r} rows x {c} columns" for r, c in sizes]
    lines.append("Headings:")
    lines += [f"  {'  ' * (h['level'] - 1)}H{h['level']} {h['text']}" for h in headings]
    lines.append("Styles used: " + ", ".join(f"{k} ({v})" for k, v in styles.items()))
    lines.append(f"Comments: {'yes' if has_comments else 'no'}")
    lines.append(f"Tracked changes: {'yes' if has_tracked else 'no'}")
    emit(args, data, lines)


def parse_outline(path: Path) -> list[tuple[str, str]]:
    try:
        raw = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        raise ToolError(1, f"Outline is not UTF-8 text: {path}") from None
    blocks: list[tuple[str, str]] = []
    for line in raw.splitlines():
        s = line.strip()
        if not s:
            continue
        if s.startswith("## "):
            blocks.append(("Heading 2", s[3:].strip()))
        elif s.startswith("# "):
            blocks.append(("Heading 1", s[2:].strip()))
        elif s.startswith("- "):
            blocks.append(("List Bullet", s[2:].strip()))
        elif _NUMBERED.fullmatch(s):
            blocks.append(("List Number", _NUMBERED.fullmatch(s).group(1).strip()))
        else:
            blocks.append(("Normal", s))
    if not blocks:
        raise ToolError(1, f"Outline has no text lines: {path}")
    return blocks


def cmd_create(args) -> None:
    src = Path(args.input)
    if not src.is_file():
        raise ToolError(1, f"Input file not found: {src}")
    blocks = parse_outline(src)
    target, _ = resolve_target(args, src, "create")
    doc = Document()
    for style, text in blocks:
        if style == "Normal":
            doc.add_paragraph(text)
        else:
            doc.add_paragraph(text, style=style)

    def verify(tmp: Path):
        again = reopen(tmp)
        got = [(style_name(p, again), para_text(p)) for p in again.element.body.iter(W + "p")]
        if got != blocks:
            raise VerifyError("the paragraphs differ from the outline")
        return {"paragraphs": len(got)}

    info = finalize(doc, target, verify)
    report_write(args, {"paragraphs": info["paragraphs"], "output": str(target)},
                 [f"Created {target} with {info['paragraphs']} paragraphs."])


def _replace_pairs(args) -> list[tuple[str, str]]:
    if args.pairs is not None:
        if args.find is not None or args.replace is not None:
            raise ToolError(1, "Give either --find and --replace, or --pairs, not both.")
        pairs_path = Path(args.pairs)
        pairs = []
        for key, value in _read_json_object(pairs_path).items():
            if not isinstance(value, str):
                raise ToolError(1, f"Could not read JSON in {pairs_path}: the replacement for {key!r} must be text.")
            pairs.append((key, value))
    else:
        if args.find is None or args.replace is None:
            raise ToolError(1, "Give --find TEXT and --replace TEXT, or --pairs FILE.json.")
        pairs = [(args.find, args.replace)]
    for find, _new in pairs:
        if find == "":
            raise ToolError(1, "The text to find must not be empty.")
    return pairs


def cmd_replace(args) -> None:
    pairs = _replace_pairs(args)
    path = Path(args.input)
    precheck(path)
    target, _in_place = resolve_target(args, path, "replace")
    doc = open_docx(path)
    roots = story_roots(doc)
    expected = count_paragraphs(roots)
    per_pair: dict[str, int] = {}
    skipped = 0
    for find, new in pairs:
        tally = replace_everywhere(roots, find, new)
        per_pair[find] = per_pair.get(find, 0) + tally.done
        skipped += tally.skipped

    def verify(tmp: Path):
        again = reopen(tmp)
        scope = story_roots(again)
        if count_paragraphs(scope) != expected:
            raise VerifyError("the paragraph count changed")
        remaining = {
            find: sum(len(find_all(para_text(p), find)) for _l, _n, root in scope for p in root.iter(W + "p"))
            for find in per_pair
        }
        return {"remaining_after": remaining}

    info = finalize(doc, target, verify)
    total = sum(per_pair.values())
    data = {
        "replacements": total,
        "per_pair": per_pair,
        "skipped_in_tracked_change": skipped,
        "remaining_after": info["remaining_after"],
        "output": str(target),
    }
    lines = [f"Wrote {target}", f"Replacements: {total}", f"Skipped in tracked changes: {skipped}"]
    lines += [f"Remaining after replace, '{k}': {v}" for k, v in info["remaining_after"].items()]
    report_write(args, data, lines, searched=True)


def cmd_fill(args) -> None:
    path = Path(args.input)
    precheck(path)
    values = _fill_values(Path(args.values))
    target, _in_place = resolve_target(args, path, "fill")
    doc = open_docx(path)
    roots = story_roots(doc)
    state = FillState()
    for _loc, _name, root in roots:
        for p in list(root.iter(W + "p")):
            _fill_paragraph(p, values, state)
    if state.unfilled and not args.allow_unfilled:
        raise ToolError(1, f"Missing values for: {', '.join(state.unfilled)}. No file written.")

    def verify(tmp: Path):
        again = reopen(tmp)
        left = count_placeholders(story_roots(again))
        if left != state.left_behind:
            raise VerifyError(f"{left} placeholders remain, expected {state.left_behind}")
        return {"placeholders_remaining": left}

    finalize(doc, target, verify)
    data = {
        "filled": state.filled,
        "per_name": state.per_name,
        "unfilled": state.unfilled,
        "skipped_in_tracked_change": state.skipped,
        "output": str(target),
    }
    lines = [f"Wrote {target}", f"Filled: {state.filled}"]
    if state.unfilled:
        lines.append(f"Left unfilled: {', '.join(state.unfilled)}")
    report_write(args, data, lines, searched=True)


def cmd_track_show(args) -> None:
    doc = open_docx(Path(args.input))
    rows = []
    for rev in scan_revisions(story_roots(doc)):
        rows.append({
            "id": rev.rid, "kind": rev.kind, "author": rev.author, "date": rev.date,
            "text": rev.text, "paragraph_index": rev.paragraph_index,
            "location": rev.location, "detail": rev.detail,
        })
    lines = [
        f"{r['id']} {r['kind']} {r['author']} {r['date'] or ''} {r['text']!r}"
        + (f" ({r['detail']})" if r["detail"] else "")
        for r in rows
    ] or ["No tracked changes found."]
    emit(args, {"revisions": rows}, lines)


def cmd_track_resolve(args) -> None:
    if args.ids and args.all:
        raise ToolError(1, "Give --id N (repeatable) or --all, not both.")
    if not args.ids and not args.all:
        raise ToolError(1, "Give --id N (repeatable) or --all.")
    path = Path(args.input)
    precheck(path)
    target, _in_place = resolve_target(args, path, "track-resolve")
    doc = open_docx(path)
    revs = scan_revisions(story_roots(doc))
    if args.all:
        selected = revs
    else:
        for rid in args.ids:
            if not any(r.rid == rid for r in revs):
                raise ToolError(1, f"No tracked change with id {rid}.")
        selected = [r for r in revs if r.rid in args.ids]
    if any(r.kind == "unsupported" for r in selected):
        raise ToolError(2, UNSUPPORTED_MSG)
    plain_total = sum(1 for r in revs if r.kind != "unsupported")
    expected_remaining = plain_total - len(selected)
    body = doc.element.body
    comments = comments_root(doc)
    had_text = _comments_with_content(body)
    for rev in selected:
        apply_revision(rev.el, rev.kind, args.mode)
    # Word drops a comment whose range loses all its text in Accept or Reject (for example a
    # comment on text that is deleted and then accepted). Do the same, so both agree.
    for cid in sorted(had_text - _comments_with_content(body)):
        _drop_comment(body, comments, cid)

    def verify(tmp: Path):
        again = reopen(tmp)
        remaining = sum(1 for r in scan_revisions(story_roots(again)) if r.kind != "unsupported")
        if remaining != expected_remaining:
            raise VerifyError(f"{remaining} tracked changes remain, expected {expected_remaining}")
        return {"remaining": remaining}

    info = finalize(doc, target, verify)
    data = {"resolved": len(selected), "remaining": info["remaining"], "output": str(target)}
    report_write(args, data, [f"Wrote {target}", f"Resolved: {len(selected)}", f"Remaining: {info['remaining']}"])


def cmd_comments_list(args) -> None:
    doc = open_docx(Path(args.input))
    rows = list_comments(doc)
    lines = [f"[{r['id']}] {r['author']} on {r['anchor_text']!r}: {r['text']}" for r in rows] or ["No comments found."]
    emit(args, {"comments": rows}, lines)


def cmd_comments_add(args) -> None:
    if not args.anchor:
        raise ToolError(1, "The anchor text must not be empty.")
    path = Path(args.input)
    precheck(path)
    target, _in_place = resolve_target(args, path, "comments-add")
    doc = open_docx(path)
    hits = []
    for p in doc.element.body.iter(W + "p"):
        for start in find_all(para_text(p), args.anchor):
            hits.append((p, start))
    if not hits:
        raise ToolError(1, "Anchor text not found in any paragraph.")
    count = len(hits)
    if args.occurrence is None:
        if count > 1:
            raise ToolError(1, f"Anchor text occurs {count} times. Pass --occurrence 1..{count}.")
        occurrence = 1
    else:
        occurrence = args.occurrence
    if not 1 <= occurrence <= count:
        raise ToolError(1, f"Anchor text occurs {count} times. Pass --occurrence 1..{count}.")
    p, start = hits[occurrence - 1]
    end = start + len(args.anchor)
    _ensure_boundary(p, start)
    _ensure_boundary(p, end)
    runs = _anchored_runs(p, start, end)
    if not runs:
        raise RuntimeError("no runs found for the anchor")
    para = Paragraph(p, doc)
    new_comment = doc.add_comment(runs=[Run(runs[0], para), Run(runs[-1], para)],
                                  text=args.text, author=args.author, initials=args.initials)
    _keep_markers_outside_revisions(doc.element.body, new_comment.comment_id)

    def verify(tmp: Path):
        again = reopen(tmp)
        match = [r for r in list_comments(again)
                 if r["anchor_text"] == args.anchor and r["text"] == args.text and r["author"] == args.author]
        if not match:
            raise VerifyError("the new comment is not found with the same anchor, text and author")
        return {"comment": match[-1]}

    info = finalize(doc, target, verify)
    comment = info["comment"]
    data = {"comment_id": comment["id"], "anchor_text": comment["anchor_text"], "text": comment["text"],
            "author": comment["author"], "output": str(target)}
    report_write(args, data, [f"Wrote {target}", f"Comment {comment['id']} on {comment['anchor_text']!r} added."])


def _check_block_names() -> list[str]:
    return ["file is a readable zip", "required parts present", "XML parts parse",
            "internal relationship targets exist", "python-docx opens the file",
            "no w:delText outside w:del", "no nested tracked changes"]


def cmd_validate(args) -> int:
    path = Path(args.input)
    precheck(path)
    rows: list[dict] = []

    def add(name: str, problems: list[str]) -> None:
        rows.append({"name": name, "passed": not problems,
                     "detail": "; ".join(problems) if problems else "ok"})

    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            data = {n: zf.read(n) for n in zf.namelist() if not n.endswith("/")}
        zip_problems = [f"file is not a readable zip (CRC error in {bad})"] if bad else []
    except PermissionError:
        raise
    except Exception as exc:
        data = {}
        zip_problems = [f"file is not a readable zip (corrupt or truncated): {exc}"]
    add("file is a readable zip", zip_problems)
    if zip_problems:
        for name in _check_block_names()[1:]:
            rows.append({"name": name, "passed": None, "detail": "skipped"})
        return _emit_validate(args, rows)

    names = set(data)
    add("required parts present",
        [f"missing {n}" for n in ("[Content_Types].xml", "word/document.xml") if n not in names])

    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    parsed, problems = {}, []
    for name, raw in data.items():
        if name.endswith((".xml", ".rels")):
            try:
                parsed[name] = etree.fromstring(raw, parser)
            except etree.XMLSyntaxError as exc:
                problems.append(f"{name}: {exc}")
    add("XML parts parse", problems)

    problems = []
    for name, root in parsed.items():
        if not name.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(name))
        for rel in root.iter(PKG_REL):
            if rel.get("TargetMode") == "External":
                continue
            target = unquote((rel.get("Target") or "").split("#")[0])
            if not target:
                continue
            if target.startswith("/"):
                full = target.lstrip("/")
            else:
                full = posixpath.normpath(posixpath.join(base, target))
            if full not in names:
                problems.append(f"{name}: target {target} not found in the file")
    add("internal relationship targets exist", problems)

    try:
        Document(str(path))
        open_problems: list[str] = []
    except PermissionError:
        raise
    except Exception as exc:
        open_problems = [f"python-docx could not open the file: {type(exc).__name__}: {exc}"]
    add("python-docx opens the file", open_problems)

    part_names = [n for n in parsed if n.startswith("word/") and n.endswith(".xml")]
    problems = []
    for name in part_names:
        if any(not any(a.tag == W + "del" for a in el.iterancestors())
               for el in parsed[name].iter(W + "delText")):
            problems.append(f"{name}: w:delText outside w:del")
    add("no w:delText outside w:del", problems)

    problems = []
    for name in part_names:
        for el in parsed[name].iter(W + "ins", W + "del"):
            if any(a.tag in (W + "ins", W + "del") for a in el.iterancestors()):
                problems.append(f"{name}: tracked change inside another tracked change")
                break
    add("no nested tracked changes", problems)
    return _emit_validate(args, rows)


def _emit_validate(args, rows: list[dict]) -> int:
    invalid = any(r["passed"] is False for r in rows)
    verdict = "invalid" if invalid else "valid"

    def label(r):
        return "SKIP" if r["passed"] is None else ("PASS" if r["passed"] else "FAIL")

    lines = [f"{label(r)} {r['name']}: {r['detail']}" for r in rows] + [verdict]
    emit(args, {"checks": rows, "verdict": verdict}, lines)
    return 1 if invalid else 0


def _find_soffice() -> str | None:
    for name in ("soffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return found
    if sys.platform == "win32":
        candidates = [r"C:\Program Files\LibreOffice\program\soffice.exe"]
    elif sys.platform == "darwin":
        candidates = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    else:
        candidates = ["/usr/bin/soffice"]
    for candidate in candidates:
        if Path(candidate).is_file():
            return candidate
    return None


def _format_seconds(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else str(value)


def _word_available() -> bool:
    """True when Microsoft Word is registered for automation and pywin32 can be imported (Windows)."""
    if sys.platform != "win32" or importlib.util.find_spec("win32com") is None:
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except OSError:
        return False


def _word_export_pdf(src: Path, dst: Path) -> None:
    """Open `src` in Word and save it as PDF (wdFormatPDF = 17). Runs only in the to-pdf child process."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        doc.SaveAs2(str(dst), FileFormat=17)
    finally:
        try:
            if doc is not None:
                doc.Close(0)  # wdDoNotSaveChanges
        except Exception:  # noqa: BLE001 - closing must not hide the export result
            pass
        try:
            if app is not None:
                app.Quit(0)
        except Exception:  # noqa: BLE001
            pass
        pythoncom.CoUninitialize()


def cmd_word_pdf(args) -> None:
    """Hidden command: the Word export in its own process, so to-pdf can stop it on timeout."""
    try:
        _word_export_pdf(Path(args.input), Path(args.out))
    except Exception as exc:  # noqa: BLE001 - every Word failure becomes one readable line
        raise ToolError(1, f"Word could not save the PDF: {type(exc).__name__}: {exc}") from None


def _word_pids() -> set[int]:
    listing = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, check=False)
    pids = set()
    for line in listing.stdout.splitlines():
        fields = line.strip().strip('"').split('","')
        if len(fields) > 1 and fields[1].isdigit():
            pids.add(int(fields[1]))
    return pids


def _convert_with_word(src: Path, produced: Path, timeout: float) -> None:
    cmd = [sys.executable, str(Path(__file__).resolve()), "_word-pdf", str(src), "--out", str(produced)]
    before = _word_pids()
    try:
        result = subprocess.run(cmd, stdin=subprocess.DEVNULL, capture_output=True, text=True,
                                encoding="utf-8", errors="replace", timeout=timeout, check=False)
    except subprocess.TimeoutExpired:
        # Killing the child leaves the Word instance it started running; stop the new ones.
        for pid in _word_pids() - before:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, check=False)
        raise ToolError(1, f"PDF conversion timed out after {_format_seconds(timeout)} seconds.") from None
    if result.returncode != 0 or not produced.is_file():
        lines = (result.stderr or "").strip().splitlines()
        raise ToolError(1, lines[-1] if lines else f"Word produced no PDF for: {src.name}")


def _convert_with_libreoffice(soffice: str, path: Path, workdir: Path, timeout: float) -> Path:
    profile_uri = (workdir / "profile").as_uri()
    cmd = [soffice, f"-env:UserInstallation={profile_uri}", "--headless", "--norestore",
           "--convert-to", "pdf", "--outdir", str(workdir), str(path.resolve())]
    with open(workdir / "soffice.log", "wb") as log:
        try:
            subprocess.run(cmd, stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT,
                           timeout=timeout, check=False)
        except subprocess.TimeoutExpired:
            raise ToolError(1, f"PDF conversion timed out after {_format_seconds(timeout)} seconds.") from None
    produced = workdir / f"{path.stem}.pdf"
    if not produced.is_file():
        pdfs = sorted(workdir.glob("*.pdf"))
        if not pdfs:
            raise ToolError(1, f"LibreOffice produced no PDF for: {path}")
        produced = pdfs[0]
    return produced


def cmd_to_pdf(args) -> None:
    path = Path(args.input)
    open_docx(path)  # missing, legacy, encrypted, macro and corrupt inputs are refused here
    out = Path(args.out) if args.out else path.with_name(f"{path.stem}.to-pdf.pdf")
    if out.suffix.lower() != ".pdf":
        raise ToolError(1, f"Output must end in .pdf: {out}")
    if out.exists() and not args.force:
        raise ToolError(1, f"Output already exists: {out}. Choose another --out or pass --force.")
    if not out.parent.is_dir():
        raise ToolError(1, f"Output folder does not exist: {out.parent}")
    if args.timeout <= 0:
        raise ToolError(1, "--timeout must be a positive number of seconds.")
    # Word is preferred when it is installed; LibreOffice is the fallback; otherwise the spec's refusal.
    word = _word_available()
    soffice = None if word else _find_soffice()
    if not word and soffice is None:
        raise ToolError(2, LIBREOFFICE_MSG)
    try:
        import pypdf
    except ImportError:
        raise ToolError(1, "pypdf is not installed. Run: pip install -r requirements.txt") from None

    converter = "Microsoft Word" if word else "LibreOffice"
    # Absolute paths matter: Word resolves relative paths from its own folder, not from ours.
    workdir = Path(tempfile.mkdtemp(prefix=".docx-tool-pdf-", dir=str(out.parent.resolve())))
    try:
        if word:
            source = workdir / path.name  # Word works on a copy: no lock file next to the input
            shutil.copy2(path, source)
            produced = workdir / f"{path.stem}.pdf"
            _convert_with_word(source, produced, args.timeout)
        else:
            produced = _convert_with_libreoffice(soffice, path, workdir, args.timeout)
        try:
            pages = len(pypdf.PdfReader(str(produced)).pages)
        except Exception as exc:
            raise ToolError(1, f"Output failed verification and was not kept: {exc}") from None
        if pages < 1:
            raise ToolError(1, "Output failed verification and was not kept: the PDF has no pages.")
        os.replace(str(produced), str(out))
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    report_write(args, {"pages": pages, "output": str(out), "converter": converter},
                 [f"Wrote {out} ({pages} pages, converted with {converter})"])


# --------------------------------------------------------------- entry point

def build_parser() -> argparse.ArgumentParser:
    parser = DocxArgumentParser(prog="docx_tool.py", description="Read and change .docx files.")
    sub = parser.add_subparsers(dest="command", metavar="command", required=True,
                                parser_class=DocxArgumentParser)

    def add(name, func, help_text):
        p = sub.add_parser(name, help=help_text)
        p.set_defaults(func=func)
        return p

    def add_output(p):
        p.add_argument("--out", help="output file path")
        p.add_argument("--force", action="store_true", help="replace an existing output")

    def add_json(p):
        p.add_argument("--json", action="store_true", help="print one JSON object")

    def add_in_place(p):
        p.add_argument("--in-place", action="store_true", help="change the input file itself")

    p = add("read", cmd_read, "print paragraphs and tables in order")
    p.add_argument("input")
    add_json(p)
    p.add_argument("--no-tables", action="store_true")

    p = add("create", cmd_create, "create a .docx from a text outline")
    p.add_argument("input", help="outline .txt file")
    add_output(p)
    add_json(p)

    p = add("replace", cmd_replace, "find and replace text")
    p.add_argument("input")
    p.add_argument("--find")
    p.add_argument("--replace")
    p.add_argument("--pairs", help="JSON object of find text to replacement text")
    add_output(p)
    add_json(p)
    add_in_place(p)

    p = add("fill", cmd_fill, "fill {{name}} placeholders from JSON")
    p.add_argument("input")
    p.add_argument("--values", required=True, help="JSON object of name to value")
    p.add_argument("--allow-unfilled", action="store_true")
    add_output(p)
    add_json(p)
    add_in_place(p)

    p = add("track-show", cmd_track_show, "list tracked changes")
    p.add_argument("input")
    add_json(p)

    p = add("track-resolve", cmd_track_resolve, "accept or reject tracked changes")
    p.add_argument("input")
    p.add_argument("--mode", required=True, choices=["accept", "reject"])
    p.add_argument("--id", dest="ids", action="append", type=int, help="revision id (repeatable)")
    p.add_argument("--all", action="store_true")
    add_output(p)
    add_json(p)
    add_in_place(p)

    p = add("comments-list", cmd_comments_list, "list comments")
    p.add_argument("input")
    add_json(p)

    p = add("comments-add", cmd_comments_add, "add a comment anchored to exact text")
    p.add_argument("input")
    p.add_argument("--anchor", required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--author", required=True)
    p.add_argument("--initials")
    p.add_argument("--occurrence", type=int)
    add_output(p)
    add_json(p)
    add_in_place(p)

    p = add("validate", cmd_validate, "check that the file opens and its XML is intact")
    p.add_argument("input")
    add_json(p)

    p = add("to-pdf", cmd_to_pdf, "convert to PDF with Microsoft Word or LibreOffice")
    p.add_argument("input")
    p.add_argument("--out")
    p.add_argument("--force", action="store_true")
    p.add_argument("--timeout", type=float, default=120.0, help="seconds (default 120)")

    # Internal: to-pdf starts this in a child process to run Word. Not listed in --help.
    p = add("_word-pdf", cmd_word_pdf, argparse.SUPPRESS)
    p.add_argument("input")
    p.add_argument("--out", required=True)

    p = add("summary", cmd_summary, "outline and counts of the document")
    p.add_argument("input")
    add_json(p)
    return parser


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main(argv=None) -> int:
    _configure_stdio()
    try:
        args = build_parser().parse_args(argv)
        result = args.func(args)
        return result if isinstance(result, int) else 0
    except ToolError as exc:
        print(exc.message, file=sys.stderr)
        return exc.code
    except PermissionError:
        print(LOCKED_MSG, file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001 - every unexpected failure ends here
        print(f"Internal error ({type(exc).__name__}). No output was kept.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
