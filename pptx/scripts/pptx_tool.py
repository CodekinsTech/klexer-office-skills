#!/usr/bin/env python3
"""pptx_tool.py - command-line tool for PowerPoint .pptx decks.

One entry point, many subcommands: create, read, replace, add, delete,
reorder, duplicate, split, replace-image, chart-data, check-text, validate.
Run `python pptx_tool.py --help` or see SKILL.md next to this folder.

Exit codes: 0 done and verified; 1 input, option or verification problem;
2 unsupported file or content (legacy, encrypted, macro-enabled, unsupported
chart or slide content).
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import os
import posixpath
import re
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.chart.plot import BarPlot, LinePlot, PiePlot
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.parts.image import Image as PptxImage
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture

VERSION = "2.0.0"
TITLE_LAYOUT = "Title and Content"
TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
BODY_TYPES = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
URI_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
URI_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
URI_OLE = "http://schemas.openxmlformats.org/presentationml/2006/ole"

LOCKED_MSG = "The file is open in another program. Close it and try again."
LEGACY_MSG = "Legacy .ppt is not supported. Save as .pptx and try again."
MACRO_MSG = "Macro-enabled presentations are not supported. Save as .pptx (macros would be dropped)."
ENCRYPTED_MSG = "This deck is password-protected. The skill does not open protected files."
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ENC_ASCII = b"EncryptedPackage"
ENC_UTF16 = "EncryptedPackage".encode("utf-16-le")
ALLOWED_IMAGE_EXT = {"png", "jpg", "gif", "bmp", "tiff"}
BULLET_RE = re.compile(r"^( *)- (.*)$")
RANGE_RE = re.compile(r"^(\d+)(?:(-)(\d*))?$")
XML_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)

NOT_SEARCHED = ["chart text", "SmartArt", "alt text", "masters and layouts"]
VALIDATE_CHECKS = [
    "readable zip",
    "content types and presentation part present",
    "XML parts parse",
    "relationship targets exist",
    "python-pptx opens the deck",
    "slide ids unique and at least 256",
    "every slide has a layout relationship",
    "notes slides point back to their slides",
    "picture relationships resolve",
]


# Features the tool does not provide (spec section 1). Asking for one is a refusal:
# exit 2 with "Not supported: <feature>. <alternative>." An unknown command name
# or a usage typo is not in this table and exits 1.
UNSUPPORTED_COMMANDS = {
    "render": ("rendering slides to PDF or images",
               "Export the deck from PowerPoint with File > Export, or use a conversion tool outside this skill."),
    "export": ("rendering slides to PDF or images",
               "Export the deck from PowerPoint with File > Export, or use a conversion tool outside this skill."),
    "pdf": ("rendering slides to PDF or images",
            "Export the deck from PowerPoint with File > Export, or use a conversion tool outside this skill."),
    "animate": ("animations", "Add or change animations in PowerPoint."),
    "transition": ("transitions", "Add or change transitions in PowerPoint."),
    "theme": ("theme editing", "Edit the theme in PowerPoint."),
    "master": ("slide master editing", "Edit the slide master in PowerPoint."),
}


class ToolError(Exception):
    """A failure with the exit code and the exact message to print on stderr."""

    def __init__(self, message: str, code: int = 1):
        super().__init__(message)
        self.message = message
        self.code = code


class VerifyError(Exception):
    """The re-opened output did not match what was intended."""


# ---------------------------------------------------------------- console

def _print(text: str = "") -> None:
    sys.stdout.write(text + "\n")


def _eprint(text: str) -> None:
    sys.stderr.write(text + "\n")


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


# ---------------------------------------------------------------- files

def check_format(path: Path) -> None:
    """Refuse legacy, encrypted and macro-enabled files (exit 2). Never unlocks or strips."""
    suffix = path.suffix.lower()
    if suffix == ".ppt":
        raise ToolError(LEGACY_MSG, 2)
    if suffix in (".pptm", ".potm"):
        raise ToolError(MACRO_MSG, 2)
    with path.open("rb") as fh:
        head = fh.read(8)
    if head == OLE_MAGIC:
        blob = path.read_bytes()
        if ENC_ASCII in blob or ENC_UTF16 in blob:
            raise ToolError(ENCRYPTED_MSG, 2)
        raise ToolError(LEGACY_MSG, 2)
    if head[:2] == b"PK":
        try:
            with zipfile.ZipFile(path) as zf:
                names = zf.namelist()
                if "ppt/vbaProject.bin" in names:
                    raise ToolError(MACRO_MSG, 2)
                if "[Content_Types].xml" in names and b"macroEnabled" in zf.read("[Content_Types].xml"):
                    raise ToolError(MACRO_MSG, 2)
        except zipfile.BadZipFile:
            pass


def open_deck(path: Path):
    if not path.is_file():
        raise ToolError(f"Input file not found: {path}", 1)
    check_format(path)
    try:
        return Presentation(str(path))
    except PermissionError:
        raise
    except Exception:
        raise ToolError(f"File is not a readable .pptx (corrupt or truncated): {path}", 1)


def reopen(path: Path):
    """Re-open a file we just wrote, for verification."""
    try:
        return Presentation(str(path))
    except Exception as err:
        raise VerifyError(f"the saved file could not be re-opened ({type(err).__name__})")


def check_out_format(out: Path) -> None:
    suffix = out.suffix.lower()
    if suffix == ".ppt":
        raise ToolError(LEGACY_MSG, 2)
    if suffix in (".pptm", ".potm"):
        raise ToolError(MACRO_MSG, 2)
    if suffix != ".pptx":
        raise ToolError(f"Output must end in .pptx: {out}", 1)


def check_destination(out: Path, force: bool) -> None:
    check_out_format(out)
    if out.exists() and not force:
        raise ToolError(f"Output already exists: {out}. Choose another --out or pass --force.", 1)
    if not out.parent.is_dir():
        raise ToolError(f"Output folder does not exist: {out.parent}", 1)


def _same_file(a: Path, b: Path) -> bool:
    try:
        return os.path.samefile(a, b)
    except OSError:
        return False


def plan_output(src: Path, args, suffix: str):
    """Return (output path, in_place flag) for a command that changes a deck."""
    in_place = bool(getattr(args, "in_place", False))
    out_arg = getattr(args, "out", None)
    if in_place:
        if out_arg:
            raise ToolError("Use --out or --in-place, not both.", 1)
        return src, True
    out = Path(out_arg) if out_arg else src.with_name(f"{src.stem}.{suffix}.pptx")
    if _same_file(out, src):
        raise ToolError("Refusing to overwrite the input. Use --in-place to change it.", 1)
    check_destination(out, bool(getattr(args, "force", False)))
    return out, False


def commit_deck(prs, out: Path, verify) -> None:
    """Save to a temp file in the output folder, verify by re-opening, then move into place."""
    fd, tmp_name = tempfile.mkstemp(prefix=".pptx-tool-", suffix=".pptx", dir=str(out.parent))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        prs.save(str(tmp))
        try:
            verify(tmp)
        except VerifyError as err:
            raise ToolError(f"Output failed verification and was not kept: {err}", 1)
        os.replace(str(tmp), str(out))
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass


def write_json_file(path: Path, obj, force: bool) -> None:
    if path.exists() and not force:
        raise ToolError(f"Output already exists: {path}. Choose another --out or pass --force.", 1)
    if not path.parent.is_dir():
        raise ToolError(f"Output folder does not exist: {path.parent}", 1)
    fd, tmp_name = tempfile.mkstemp(prefix=".pptx-tool-", suffix=".json", dir=str(path.parent))
    os.close(fd)
    try:
        Path(tmp_name).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp_name, str(path))
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def read_json_file(path: Path):
    if not path.is_file():
        raise ToolError(f"Input file not found: {path}", 1)
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as err:
        raise ToolError(f"Could not read {path}: {err}", 1)


# ---------------------------------------------------------------- slide helpers

def all_layouts(prs) -> list:
    layouts = []
    for master in prs.slide_masters:
        layouts.extend(master.slide_layouts)
    return layouts


def find_placeholder(container, types):
    for ph in container.placeholders:
        if ph.placeholder_format.type in types:
            return ph
    return None


def is_usable_layout(layout) -> bool:
    return find_placeholder(layout, TITLE_TYPES) is not None and find_placeholder(layout, BODY_TYPES) is not None


def find_layout(prs, name: str):
    layouts = all_layouts(prs)
    for layout in layouts:
        if layout.name == name:
            return layout
    names = []
    for layout in layouts:
        if layout.name not in names:
            names.append(layout.name)
    raise ToolError(f'No layout named "{name}" in the template. Layouts are: {", ".join(names)}.', 1)


def iter_shapes(shapes):
    for shape in shapes:
        if isinstance(shape, GroupShape):
            yield from iter_shapes(shape.shapes)
        else:
            yield shape


def slide_title(slide) -> str:
    title = slide.shapes.title
    if title is None or not title.has_text_frame:
        return ""
    return title.text_frame.text


def read_notes(slide) -> str:
    if not slide.has_notes_slide:
        return ""
    tf = slide.notes_slide.notes_text_frame
    return tf.text if tf is not None else ""


def slide_body_lines(slide) -> list:
    ph = find_placeholder(slide, BODY_TYPES)
    if ph is None or not ph.has_text_frame:
        return []
    return [(p.level + 1, p.text) for p in ph.text_frame.paragraphs if p.text.strip()]


def clean_number(value):
    if value is None:
        return None
    number = float(value)
    if math.isfinite(number) and number.is_integer():
        return int(number)
    return number


def chart_type_name(chart) -> str:
    try:
        return chart.chart_type.name
    except Exception:
        return "UNKNOWN"


def chart_record(chart) -> dict:
    categories = []
    for plot in chart.plots:
        try:
            labels = [str(c) for c in plot.categories]
        except Exception:
            labels = []
        if labels:
            categories = labels
            break
    series = [
        {"name": s.name, "values": [clean_number(v) for v in s.values]}
        for s in chart.series
    ]
    return {"type": chart_type_name(chart), "categories": categories, "series": series}


def slide_contents(slide):
    """Return (texts, tables, charts) for one slide, groups expanded."""
    texts, tables, charts = [], [], []
    for shape in iter_shapes(slide.shapes):
        if shape.has_text_frame:
            text = shape.text_frame.text
            if text.strip():
                texts.append(text)
        if shape.has_table:
            tables.append([[cell.text for cell in row.cells] for row in shape.table.rows])
        if shape.has_chart:
            charts.append(chart_record(shape.chart))
    return texts, tables, charts


def picture_shapes(slide) -> list:
    return [sh for sh in slide.shapes if isinstance(sh, Picture)]


def picture_hashes(slide) -> list:
    return [hashlib.sha256(pic.image.blob).hexdigest() for pic in picture_shapes(slide)]


def unsupported_kind(slide):
    """Name the content that duplicate cannot copy completely, or None."""
    for node in slide._element.iter():
        if not isinstance(node.tag, str):
            continue
        if node.tag == qn("p:oleObj"):
            return "OLE object"
        if node.tag == qn("a:graphicData"):
            uri = node.get("uri", "")
            if uri == URI_CHART:
                return "chart"
            if uri == URI_DIAGRAM:
                return "SmartArt"
            if uri == URI_OLE:
                return "OLE object"
        if node.tag in (qn("a:videoFile"), qn("a:audioFile"), qn("a:wavAudioFile"), f"{{{NS_P14}}}media"):
            return "video or audio"
    return None


def move_slide(prs, from_index: int, to_index: int) -> None:
    lst = prs.slides._sldIdLst
    element = list(lst)[from_index]
    lst.remove(element)
    lst.insert(to_index, element)


def delete_slide_at(prs, index: int) -> None:
    """Remove one slide from the slide list and drop its relationship (python-pptx omits unreachable parts on save)."""
    lst = prs.slides._sldIdLst
    sld_id = list(lst)[index]
    rid = sld_id.get(qn("r:id"))
    slide_id = sld_id.get("id")
    lst.remove(sld_id)
    pres = prs.part._element
    for node in [n for n in pres.iter(f"{{{NS_P14}}}sldId") if n.get("id") == slide_id]:
        node.getparent().remove(node)
    for node in [n for n in pres.iter(f"{{{NS_P}}}sld") if n.get(qn("r:id")) == rid]:
        node.getparent().remove(node)
    prs.part.drop_rel(rid)
    prs.part.rename_slide_parts([e.get(qn("r:id")) for e in lst])


def remap_relationships(element, src_part, dst_part) -> None:
    """Give every relationship id inside a copied element a matching relationship on the destination part."""
    attr_prefix = f"{{{NS_R}}}"
    for node in element.iter():
        if not isinstance(node.tag, str):
            continue
        for key, value in list(node.attrib.items()):
            if not key.startswith(attr_prefix):
                continue
            try:
                rel = src_part.rels[value]
            except KeyError:
                continue
            if rel.is_external:
                new_id = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                new_id = dst_part.relate_to(rel.target_part, rel.reltype)
            node.set(key, new_id)


def rid_in_use(root, rid: str) -> bool:
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        if rid in node.attrib.values():
            return True
    return False


# ---------------------------------------------------------------- slide lists

def _out_of_range(number: int, count: int) -> ToolError:
    return ToolError(f"Slide {number} does not exist. The deck has {count} slides.", 1)


def parse_ranges(text: str, count: int) -> list:
    """Parse '1-3,5,8-' into (first, last) pairs in the order given."""
    invalid = ToolError(f"Slide list is invalid: {text}.", 1)
    if text is None or not text.strip():
        raise invalid
    ranges = []
    for item in text.split(","):
        match = RANGE_RE.match(item.strip())
        if not match:
            raise invalid
        first = int(match.group(1))
        if match.group(2) is None:
            last = first
        elif match.group(3) == "":
            last = count
        else:
            last = int(match.group(3))
        if first < 1 or last < first:
            raise invalid
        if first > count:
            raise _out_of_range(first, count)
        if last > count:
            raise _out_of_range(last, count)
        ranges.append((first, last))
    return ranges


def parse_slide_list(text: str, count: int, keep_order: bool = False) -> list:
    numbers = [k for first, last in parse_ranges(text, count) for k in range(first, last + 1)]
    return numbers if keep_order else sorted(set(numbers))


# ---------------------------------------------------------------- outline and create

def parse_body_line(raw: str):
    """'- text' is level 1, two spaces then '- ' is level 2, other text is level 1."""
    match = BULLET_RE.match(raw.rstrip())
    if match:
        level = 2 if len(match.group(1)) >= 2 else 1
        return level, match.group(2).strip()
    return 1, raw.strip()


def parse_outline(text: str) -> list:
    slides = []
    current = None
    for number, raw in enumerate(text.splitlines(), start=1):
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("# "):
            current = {"title": line[2:].strip(), "body": [], "notes": []}
            slides.append(current)
            continue
        if current is None:
            raise ToolError(f'Line {number} of the outline has text before the first "# " title.', 1)
        stripped = line.strip()
        if stripped.startswith("Notes:"):
            note = stripped[len("Notes:"):].strip()
            if note:
                current["notes"].append(note)
            continue
        level, body_text = parse_body_line(line)
        if body_text:
            current["body"].append((level, body_text))
    if not slides:
        raise ToolError("The outline has no slides.", 1)
    return slides


def fill_body(placeholder, lines) -> None:
    tf = placeholder.text_frame
    for index, (level, text) in enumerate(lines):
        paragraph = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        paragraph.text = text
        paragraph.level = level - 1


def remove_empty_placeholders(slide) -> None:
    for ph in list(slide.placeholders):
        if ph.has_text_frame and not ph.text_frame.text.strip():
            element = ph._element
            element.getparent().remove(element)


def populate_slide(slide, title, body, notes) -> None:
    title_ph = slide.shapes.title
    if title is not None and title_ph is not None:
        title_ph.text_frame.text = title
    body_ph = find_placeholder(slide, BODY_TYPES)
    if body:
        if body_ph is None:
            raise ToolError(f'Layout "{slide.slide_layout.name}" has no body placeholder for the body lines.', 1)
        fill_body(body_ph, body)
    if notes:
        slide.notes_slide.notes_text_frame.text = "\n".join(notes)
    remove_empty_placeholders(slide)


def choose_create_layout(prs):
    layouts = all_layouts(prs)
    for layout in layouts:
        if layout.name == TITLE_LAYOUT:
            return layout, None
    for layout in layouts:
        if is_usable_layout(layout):
            return layout, f'"{TITLE_LAYOUT}" was not found, so layout "{layout.name}" was used.'
    raise ToolError("The template has no usable layout. A layout needs a title and a body placeholder.", 1)


def verify_outline(path: Path, spec: list) -> None:
    prs = reopen(path)
    slides = list(prs.slides)
    if len(slides) != len(spec):
        raise VerifyError(f"expected {len(spec)} slides, found {len(slides)}")
    for index, (slide, item) in enumerate(zip(slides, spec), start=1):
        if slide_title(slide) != item["title"]:
            raise VerifyError(f"slide {index} title is not {item['title']!r}")
        if slide_body_lines(slide) != item["body"]:
            raise VerifyError(f"slide {index} bullets do not match the outline")
        if read_notes(slide) != "\n".join(item["notes"]):
            raise VerifyError(f"slide {index} notes do not match the outline")


def cmd_create(args) -> int:
    src = Path(args.outline)
    if not src.is_file():
        raise ToolError(f"Input file not found: {src}", 1)
    try:
        text = src.read_bytes().decode("utf-8-sig")
    except UnicodeDecodeError:
        raise ToolError(f"Outline is not UTF-8 text: {src}", 1)
    spec = parse_outline(text)
    out = Path(args.out)
    check_destination(out, args.force)
    template = Path(args.template) if args.template else None
    if template is not None:
        prs = open_deck(template)
        while len(prs.slides) > 0:
            delete_slide_at(prs, 0)
    else:
        prs = Presentation()
    layout, note = choose_create_layout(prs)
    for item in spec:
        slide = prs.slides.add_slide(layout)
        populate_slide(slide, item["title"], item["body"], item["notes"])
    commit_deck(prs, out, lambda p: verify_outline(p, spec))
    result = {"output": str(out), "slides": len(spec), "layout": layout.name,
              "layout_note": note, "verified": True}
    lines = [f"Wrote {out} with {len(spec)} slides."]
    if note:
        lines.append(note)
    emit_action(args, result, lines)
    return 0


# ---------------------------------------------------------------- output helpers

def emit_action(args, result: dict, lines: list) -> None:
    if args.json:
        _print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for line in lines:
            _print(line)


def emit_read(args, result: dict, render) -> None:
    if getattr(args, "out", None):
        write_json_file(Path(args.out), result, args.force)
    if args.json:
        _print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        render(result)
        if getattr(args, "out", None):
            _print(f"Wrote {args.out}")


# ---------------------------------------------------------------- read

def slide_record(slide, number: int) -> dict:
    texts, tables, charts = slide_contents(slide)
    return {
        "number": number,
        "layout": slide.slide_layout.name,
        "title": slide_title(slide),
        "texts": texts,
        "tables": tables,
        "charts": charts,
        "notes": read_notes(slide),
    }


def render_read(result: dict) -> None:
    _print(f"Slides: {result['slide_count']}")
    for item in result["slides"]:
        _print(f"--- slide {item['number']} ---")
        _print(f"Layout: {item['layout']}")
        _print(f"Title: {item['title']}")
        for text in item["texts"]:
            _print(f"Text: {text}")
        for index, table in enumerate(item["tables"], start=1):
            _print(f"Table {index}:")
            for row in table:
                _print("  | " + " | ".join(row) + " |")
        for chart in item["charts"]:
            _print(f"Chart: {chart['type']}")
            _print(f"  Categories: {', '.join(chart['categories'])}")
            for series in chart["series"]:
                _print(f"  Series {series['name']}: {series['values']}")
        _print("Notes:")
        if item["notes"]:
            _print(item["notes"])


def cmd_read(args) -> int:
    prs = open_deck(Path(args.deck))
    slides = list(prs.slides)
    count = len(slides)
    numbers = parse_slide_list(args.slides, count) if args.slides else list(range(1, count + 1))
    result = {
        "slide_count": count,
        "slides": [slide_record(slides[k - 1], k) for k in numbers],
    }
    emit_read(args, result, render_read)
    return 0


# ---------------------------------------------------------------- replace

def para_segments(p_el):
    """Split a paragraph into (run element or None, text) pieces. Breaks and fields are not text."""
    segments = []
    for child in p_el:
        if child.tag == qn("a:r"):
            t = child.find(qn("a:t"))
            segments.append((child, (t.text or "") if t is not None else ""))
        elif child.tag == qn("a:br"):
            segments.append((None, "\v"))
        elif child.tag == qn("a:fld"):
            segments.append((None, "\x00"))
    return segments


def run_flags(segments) -> list:
    flags = []
    for element, text in segments:
        flags.extend([element is not None] * len(text))
    return flags


def count_in_paragraph(p_el, needle: str) -> int:
    segments = para_segments(p_el)
    full = "".join(text for _, text in segments)
    flags = run_flags(segments)
    count, pos = 0, 0
    while True:
        start = full.find(needle, pos)
        if start < 0:
            break
        end = start + len(needle)
        if all(flags[start:end]):
            count += 1
            pos = end
        else:
            pos = start + 1
    return count


def _set_run_text(element, text: str) -> None:
    element.find(qn("a:t")).text = text


def _replace_once(segments, start: int, end: int, repl: str) -> None:
    spans = []
    offset = 0
    for index, (element, text) in enumerate(segments):
        seg_start, seg_end = offset, offset + len(text)
        offset = seg_end
        if text and seg_start < end and seg_end > start:
            spans.append((index, seg_start, seg_end))
    for position, (index, seg_start, seg_end) in enumerate(spans):
        element, text = segments[index]
        if position == 0:
            head = text[: max(0, start - seg_start)]
            tail = text[end - seg_start:] if end < seg_end else ""
            _set_run_text(element, head + repl + tail)
        elif end < seg_end:
            _set_run_text(element, text[end - seg_start:])
        else:
            element.getparent().remove(element)


def replace_in_paragraph(p_el, find: str, repl: str) -> int:
    count, pos = 0, 0
    while True:
        segments = para_segments(p_el)
        full = "".join(text for _, text in segments)
        flags = run_flags(segments)
        start = full.find(find, pos)
        if start < 0:
            break
        end = start + len(find)
        if not all(flags[start:end]):
            pos = start + 1
            continue
        _replace_once(segments, start, end, repl)
        count += 1
        pos = start + len(repl)
    return count


def paragraph_scope(prs, include_notes: bool) -> list:
    items = []
    for number, slide in enumerate(prs.slides, start=1):
        for shape in iter_shapes(slide.shapes):
            if shape.has_text_frame:
                items.extend((number, False, p._p) for p in shape.text_frame.paragraphs)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        items.extend((number, False, p._p) for p in cell.text_frame.paragraphs)
        if include_notes and slide.has_notes_slide:
            tf = slide.notes_slide.notes_text_frame
            if tf is not None:
                items.extend((number, True, p._p) for p in tf.paragraphs)
    return items


def count_total(items, needle: str) -> int:
    return sum(count_in_paragraph(p_el, needle) for _, _, p_el in items)


def replacement_pairs(args) -> list:
    if args.pairs is not None:
        if args.find is not None or args.replace is not None:
            raise ToolError("Use either --find and --replace, or --pairs, not both.", 1)
        data = read_json_file(Path(args.pairs))
        if not isinstance(data, dict) or not all(isinstance(v, str) for v in data.values()):
            raise ToolError(f"Could not read {args.pairs}: expected an object that maps text to replacement text.", 1)
        pairs = list(data.items())
    elif args.find is not None and args.replace is not None:
        pairs = [(args.find, args.replace)]
    else:
        raise ToolError("Give --find and --replace, or --pairs FILE.json.", 1)
    for find, _ in pairs:
        if find == "":
            raise ToolError("The text to find must not be empty.", 1)
    return pairs


def cmd_replace(args) -> int:
    pairs = replacement_pairs(args)
    src = Path(args.deck)
    prs = open_deck(src)
    out, _ = plan_output(src, args, "replace")
    include = bool(args.include_notes)
    items = paragraph_scope(prs, include)
    single = pairs[0] if len(pairs) == 1 else None
    repl_is_new = single is not None and bool(single[1]) and single[0] not in single[1]
    before_repl = count_total(items, single[1]) if repl_is_new else None

    per_slide: dict = {}
    notes_hits = 0
    total = 0
    pair_hits = []
    for find, repl in pairs:
        hits_here = 0
        for number, is_notes, p_el in items:
            hits = replace_in_paragraph(p_el, find, repl)
            if hits:
                per_slide[number] = per_slide.get(number, 0) + hits
                total += hits
                hits_here += hits
                if is_notes:
                    notes_hits += hits
        pair_hits.append(hits_here)

    # In-memory counts after all pairs ran (pairs apply in order, so a later find can match an earlier replacement).
    remaining = {find: count_total(items, find) for find, _ in pairs}

    def verify(path: Path) -> None:
        check_items = paragraph_scope(reopen(path), include)
        for find, _ in pairs:
            found = count_total(check_items, find)
            if found != remaining[find]:
                raise VerifyError(f'"{find}" found {found} time(s) in the file, expected {remaining[find]}')
        if single is not None:
            # Single pair: what is left is only the find text that the replacement itself contains.
            left = total * single[1].count(single[0])
            if remaining[single[0]] != left:
                raise VerifyError(f'"{single[0]}" still found {remaining[single[0]] - left} time(s) after replacement')
        if before_repl is not None:
            expected = before_repl + total
            after = count_total(check_items, single[1])
            if after != expected:
                raise VerifyError(f'"{single[1]}" appears {after} time(s), expected {expected}')

    commit_deck(prs, out, verify)
    if repl_is_new:
        note = "checked: find text gone and replacement count increased by the replacement count"
    else:
        note = "checked: find-text remainder only (replacement contains the find text)"
    result = {
        "output": str(out),
        "replacements": total,
        "per_slide": {str(k): v for k, v in sorted(per_slide.items())},
        "notes_replacements": notes_hits,
        "not_searched": NOT_SEARCHED + ([] if include else ["speaker notes"]),
        "verification": note,
        "verified": True,
    }
    emit_action(args, result, [
        f"Wrote {out}",
        f"Replacements: {total} (in speaker notes: {notes_hits})",
        "Not searched: " + ", ".join(result["not_searched"]),
        f"Verification: {note}",
    ])
    return 0


# ---------------------------------------------------------------- add

def cmd_add(args) -> int:
    src = Path(args.deck)
    prs = open_deck(src)
    out, _ = plan_output(src, args, "add")
    count = len(prs.slides)
    layout = find_layout(prs, args.layout)
    at = args.at if args.at is not None else count + 1
    if not 1 <= at <= count + 1:
        raise ToolError(f"Position must be between 1 and {count + 1}. The deck has {count} slides.", 1)
    body = [(level, text) for level, text in (parse_body_line(x) for x in (args.body or [])) if text]
    if body and find_placeholder(layout, BODY_TYPES) is None:
        raise ToolError(f'Layout "{layout.name}" has no body placeholder for --body lines.', 1)
    if args.title is not None and find_placeholder(layout, TITLE_TYPES) is None:
        raise ToolError(f'Layout "{layout.name}" has no title placeholder for --title.', 1)
    notes = [args.notes] if args.notes else []
    slide = prs.slides.add_slide(layout)
    populate_slide(slide, args.title, body, notes)
    move_slide(prs, count, at - 1)

    def verify(path: Path) -> None:
        vp = reopen(path)
        slides = list(vp.slides)
        if len(slides) != count + 1:
            raise VerifyError(f"expected {count + 1} slides, found {len(slides)}")
        new = slides[at - 1]
        if args.title is not None and slide_title(new) != args.title:
            raise VerifyError("the new slide title does not match --title")
        if body and slide_body_lines(new) != body:
            raise VerifyError("the new slide bullets do not match --body")
        if notes and read_notes(new) != args.notes:
            raise VerifyError("the new slide notes do not match --notes")

    commit_deck(prs, out, verify)
    result = {"output": str(out), "slide_count": count + 1, "position": at, "verified": True}
    emit_action(args, result, [f"Wrote {out}", f"Added slide {at}; deck now has {count + 1} slides."])
    return 0


# ---------------------------------------------------------------- delete / reorder

def cmd_delete(args) -> int:
    src = Path(args.deck)
    prs = open_deck(src)
    count = len(prs.slides)
    numbers = parse_slide_list(args.slides, count)
    if len(numbers) >= count:
        raise ToolError("A deck must keep at least one slide.", 1)
    out, _ = plan_output(src, args, "delete")
    titles = [slide_title(s) for s in prs.slides]
    removed = set(numbers)
    expected = [titles[k - 1] for k in range(1, count + 1) if k not in removed]
    for number in sorted(removed, reverse=True):
        delete_slide_at(prs, number - 1)

    def verify(path: Path) -> None:
        vp = reopen(path)
        slides = list(vp.slides)
        if len(slides) != len(expected):
            raise VerifyError(f"expected {len(expected)} slides, found {len(slides)}")
        if [slide_title(s) for s in slides] != expected:
            raise VerifyError("remaining slide titles changed")

    commit_deck(prs, out, verify)
    result = {"output": str(out), "deleted": numbers, "slide_count": len(expected), "verified": True}
    emit_action(args, result, [f"Wrote {out}", f"Deleted {len(numbers)} slide(s); {len(expected)} remain."])
    return 0


def cmd_reorder(args) -> int:
    src = Path(args.deck)
    prs = open_deck(src)
    count = len(prs.slides)
    order = parse_slide_list(args.order, count, keep_order=True)
    if sorted(order) != list(range(1, count + 1)):
        raise ToolError(f"Order must list every slide 1..{count} exactly once.", 1)
    out, _ = plan_output(src, args, "reorder")
    titles = [slide_title(s) for s in prs.slides]
    expected = [titles[k - 1] for k in order]
    lst = prs.slides._sldIdLst
    elements = list(lst)
    for element in elements:
        lst.remove(element)
    for number in order:
        lst.append(elements[number - 1])

    def verify(path: Path) -> None:
        slides = list(reopen(path).slides)
        if [slide_title(s) for s in slides] != expected:
            raise VerifyError("slide titles are not in the requested order")

    commit_deck(prs, out, verify)
    result = {"output": str(out), "order": order, "verified": True}
    emit_action(args, result, [f"Wrote {out}", "Order: " + ",".join(str(k) for k in order)])
    return 0


# ---------------------------------------------------------------- duplicate

def cmd_duplicate(args) -> int:
    src = Path(args.deck)
    prs = open_deck(src)
    count = len(prs.slides)
    number = args.slide
    if not 1 <= number <= count:
        raise _out_of_range(number, count)
    source = list(prs.slides)[number - 1]
    kind = unsupported_kind(source)
    if kind:
        raise ToolError(f"Slide {number} contains a {kind}. Duplicating it is not supported.", 2)
    out, _ = plan_output(src, args, "duplicate")
    src_texts, src_tables, _ = slide_contents(source)
    src_hashes = picture_hashes(source)
    src_title = slide_title(source)

    new = prs.slides.add_slide(source.slide_layout)
    dst_tree = new.shapes._spTree
    keep = (qn("p:nvGrpSpPr"), qn("p:grpSpPr"))
    for child in list(dst_tree):
        if child.tag not in keep:
            dst_tree.remove(child)
    for child in source.shapes._spTree:
        if not isinstance(child.tag, str) or child.tag in keep:
            continue
        copied = copy.deepcopy(child)
        remap_relationships(copied, source.part, new.part)
        dst_tree.append(copied)

    src_cSld = source._element.find(qn("p:cSld"))
    bg = src_cSld.find(qn("p:bg")) if src_cSld is not None else None
    if bg is not None:
        dst_cSld = new._element.find(qn("p:cSld"))
        old_bg = dst_cSld.find(qn("p:bg"))
        if old_bg is not None:
            dst_cSld.remove(old_bg)
        copied_bg = copy.deepcopy(bg)
        remap_relationships(copied_bg, source.part, new.part)
        dst_cSld.insert(0, copied_bg)

    if source.has_notes_slide:
        notes_tf = source.notes_slide.notes_text_frame
        if notes_tf is not None and notes_tf.text:
            new.notes_slide.notes_text_frame.text = notes_tf.text

    move_slide(prs, count, number)

    def verify(path: Path) -> None:
        slides = list(reopen(path).slides)
        if len(slides) != count + 1:
            raise VerifyError(f"expected {count + 1} slides, found {len(slides)}")
        copy_slide = slides[number]
        texts, tables, _ = slide_contents(copy_slide)
        if texts != src_texts or tables != src_tables:
            raise VerifyError("the new slide text or table text differs from the source slide")
        if picture_hashes(copy_slide) != src_hashes:
            raise VerifyError("the new slide pictures differ from the source slide")
        if slide_title(copy_slide) != src_title:
            raise VerifyError("the new slide title differs from the source slide")

    commit_deck(prs, out, verify)
    result = {"output": str(out), "source_slide": number, "new_slide": number + 1, "verified": True}
    emit_action(args, result, [f"Wrote {out}", f"Slide {number} copied to slide {number + 1}."])
    return 0


# ---------------------------------------------------------------- split

def check_coverage(ranges: list, count: int) -> None:
    hits = [0] * (count + 1)
    for first, last in ranges:
        for k in range(first, last + 1):
            hits[k] += 1
    for k in range(1, count + 1):
        if hits[k] == 0:
            raise ToolError(f"Ranges must cover each slide once. Slide {k} is not covered.", 1)
    for k in range(1, count + 1):
        if hits[k] > 1:
            raise ToolError(f"Ranges must cover each slide once. Slide {k} is covered more than once.", 1)


def cmd_split(args) -> int:
    src = Path(args.deck)
    prs = open_deck(src)
    count = len(prs.slides)
    if (args.every is None) == (args.ranges is None):
        raise ToolError("Give exactly one of --every or --ranges.", 1)
    if count == 0:
        raise ToolError("The deck has no slides to split.", 1)
    if args.every is not None:
        if args.every < 1:
            raise ToolError("--every must be a whole number of 1 or more.", 1)
        ranges = [(a, min(a + args.every - 1, count)) for a in range(1, count + 1, args.every)]
    else:
        ranges = parse_ranges(args.ranges, count)
        check_coverage(ranges, count)
    titles = [slide_title(s) for s in prs.slides]

    out_dir = Path(args.out_dir) if args.out_dir else src.parent / f"{src.stem}.split"
    created = False
    if out_dir.exists():
        if not out_dir.is_dir():
            raise ToolError(f"Output path is not a folder: {out_dir}", 1)
        if any(out_dir.iterdir()):
            raise ToolError(f"Output folder is not empty: {out_dir}", 1)
    else:
        if not out_dir.parent.is_dir():
            raise ToolError(f"Output folder does not exist: {out_dir.parent}", 1)
        out_dir.mkdir()
        created = True

    staged = []  # (temp path, final path, first, last)
    try:
        for index, (first, last) in enumerate(ranges, start=1):
            part = open_deck(src)
            for k in range(len(part.slides), 0, -1):
                if not first <= k <= last:
                    delete_slide_at(part, k - 1)
            fd, tmp_name = tempfile.mkstemp(prefix=".pptx-tool-", suffix=".pptx", dir=str(out_dir))
            os.close(fd)
            tmp = Path(tmp_name)
            staged.append((tmp, out_dir / f"{src.stem}.part{index:02d}.pptx", first, last))
            part.save(str(tmp))
            expected = titles[first - 1:last]
            try:
                slides = list(reopen(tmp).slides)
                if len(slides) != last - first + 1:
                    raise VerifyError(f"part {index} has {len(slides)} slides")
                if [slide_title(s) for s in slides] != expected:
                    raise VerifyError(f"part {index} titles do not match the source slides")
            except VerifyError as err:
                raise ToolError(f"Output failed verification and was not kept: {err}", 1)
        for tmp, final, _, _ in staged:
            os.replace(str(tmp), str(final))
    except BaseException:
        for tmp, final, _, _ in staged:
            for path in (tmp, final):
                if path.exists():
                    try:
                        path.unlink()
                    except OSError:
                        pass
        if created:
            try:
                out_dir.rmdir()
            except OSError:
                pass
        raise

    parts = [
        {"file": str(final), "first_slide": first, "last_slide": last, "slides": last - first + 1}
        for _, final, first, last in staged
    ]
    result = {"out_dir": str(out_dir), "parts": parts, "verified": True}
    lines = [f"Wrote {len(parts)} part(s) to {out_dir}"]
    lines += [f"  {p['file']}: slides {p['first_slide']}-{p['last_slide']}" for p in parts]
    emit_action(args, result, lines)
    return 0


# ---------------------------------------------------------------- replace-image

def check_image(path: Path) -> bytes:
    try:
        if not path.is_file():
            raise ValueError("missing")
        image = PptxImage.from_file(str(path))
        if image.ext not in ALLOWED_IMAGE_EXT:
            raise ValueError("unsupported")
        return image.blob
    except Exception:
        raise ToolError(f"Image must be PNG, JPEG, GIF, BMP or TIFF: {path}", 1)


def cmd_replace_image(args) -> int:
    src = Path(args.deck)
    prs = open_deck(src)
    count = len(prs.slides)
    number = args.slide
    if not 1 <= number <= count:
        raise _out_of_range(number, count)
    if args.shape is not None and args.shape_index is not None:
        raise ToolError("Give --shape or --shape-index, not both.", 1)
    image_path = Path(args.image)
    image_blob = check_image(image_path)
    out, _ = plan_output(src, args, "replace-image")
    slide = list(prs.slides)[number - 1]
    pictures = picture_shapes(slide)
    if args.shape is not None:
        index = next((i for i, p in enumerate(pictures) if p.name == args.shape), None)
        if index is None:
            raise ToolError(f'Shape "{args.shape}" on slide {number} is not a picture (or does not exist).', 1)
    elif args.shape_index is not None:
        index = args.shape_index - 1
        if not 0 <= index < len(pictures):
            raise ToolError(f'Shape "#{args.shape_index}" on slide {number} is not a picture (or does not exist).', 1)
    elif len(pictures) == 1:
        index = 0
    else:
        raise ToolError(f"Give --shape NAME or --shape-index K. Slide {number} has {len(pictures)} pictures.", 1)

    old = pictures[index]
    old_el = old._element
    left, top, width, height = old.left, old.top, old.width, old.height
    cropped = any((old.crop_left, old.crop_right, old.crop_top, old.crop_bottom))
    old_rid = old_el.blip_rId
    old_name = old.name
    old_cnv = old_el.find(qn("p:nvPicPr")).find(qn("p:cNvPr"))
    old_descr = old_cnv.get("descr")
    expected_hash = hashlib.sha256(image_blob).hexdigest()

    new_pic = slide.shapes.add_picture(str(image_path), left, top, width, height)
    new_el = new_pic._element
    new_cnv = new_el.find(qn("p:nvPicPr")).find(qn("p:cNvPr"))
    new_cnv.set("name", old_name)
    if old_descr is not None:
        new_cnv.set("descr", old_descr)
    old_el.addprevious(new_el)
    old_el.getparent().remove(old_el)
    if old_rid and not rid_in_use(slide._element, old_rid):
        slide.part.rels.pop(old_rid)

    def verify(path: Path) -> None:
        vp = reopen(path)
        vslides = list(vp.slides)
        if len(vslides) != count:
            raise VerifyError("the slide count changed")
        vpics = picture_shapes(vslides[number - 1])
        if len(vpics) != len(pictures):
            raise VerifyError("the number of pictures on the slide changed")
        vpic = vpics[index]
        if hashlib.sha256(vpic.image.blob).hexdigest() != expected_hash:
            raise VerifyError("the picture does not carry the new image")
        if (vpic.left, vpic.top, vpic.width, vpic.height) != (left, top, width, height):
            raise VerifyError("the picture position or size changed")

    commit_deck(prs, out, verify)
    result = {"output": str(out), "slide": number, "shape": old_name,
              "cropped_dropped": cropped, "verified": True}
    lines = [f"Wrote {out}", f'Replaced picture "{old_name}" on slide {number}.']
    if cropped:
        lines.append("The old picture was cropped; cropping was not carried over.")
    emit_action(args, result, lines)
    return 0


# ---------------------------------------------------------------- chart-data

def load_chart_data(path: Path):
    data = read_json_file(path)
    if not isinstance(data, dict) or not isinstance(data.get("categories"), list) \
            or not isinstance(data.get("series"), list):
        raise ToolError(f'Could not read {path}: expected an object with "categories" and "series" lists.', 1)
    categories = data["categories"]
    if not categories:
        raise ToolError(f"Could not read {path}: categories must not be empty.", 1)
    for category in categories:
        if category is None or isinstance(category, (dict, list)):
            raise ToolError(f"Could not read {path}: each category must be text or a number.", 1)
    category_text = [c if isinstance(c, str) else str(c) for c in categories]
    if not data["series"]:
        raise ToolError(f"Could not read {path}: series must not be empty.", 1)
    series = []
    for item in data["series"]:
        if not isinstance(item, dict) or not isinstance(item.get("name"), str) \
                or not isinstance(item.get("values"), list):
            raise ToolError(f'Could not read {path}: each series needs a "name" text and a "values" list.', 1)
        name, values = item["name"], item["values"]
        if len(values) != len(categories):
            raise ToolError(f'Series "{name}" has {len(values)} values for {len(categories)} categories.', 1)
        cleaned = []
        for value in values:
            if value is None:
                cleaned.append(None)
            elif isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
                cleaned.append(value)
            else:
                raise ToolError(f'Chart value "{value}" is not a number or null.', 1)
        series.append((name, cleaned))
    return category_text, series


def cmd_chart_data(args) -> int:
    src = Path(args.deck)
    prs = open_deck(src)
    count = len(prs.slides)
    number = args.slide
    if not 1 <= number <= count:
        raise _out_of_range(number, count)
    categories, series = load_chart_data(Path(args.data))
    out, _ = plan_output(src, args, "chart-data")
    slide = list(prs.slides)[number - 1]
    charts = [sh for sh in iter_shapes(slide.shapes) if sh.has_chart]
    if args.shape is not None:
        index = next((i for i, c in enumerate(charts) if c.name == args.shape), None)
        if index is None:
            raise ToolError(f"No chart on slide {number}.", 1)
    elif not charts:
        raise ToolError(f"No chart on slide {number}.", 1)
    elif len(charts) > 1:
        raise ToolError(f"Slide {number} has {len(charts)} charts. Pass --shape NAME.", 1)
    else:
        index = 0
    chart = charts[index].chart
    plots = list(chart.plots)
    if not plots or not all(isinstance(p, (BarPlot, LinePlot, PiePlot)) for p in plots):
        raise ToolError("Only category charts (bar, column, line, pie) can be edited.", 2)

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series:
        chart_data.add_series(name, values)
    chart.replace_data(chart_data)
    expected_names = [name for name, _ in series]
    expected_values = [[clean_number(v) for v in values] for _, values in series]

    def verify(path: Path) -> None:
        vp = reopen(path)
        vslides = list(vp.slides)
        vcharts = [sh for sh in iter_shapes(vslides[number - 1].shapes) if sh.has_chart]
        if index >= len(vcharts):
            raise VerifyError("the chart is missing after the update")
        record = chart_record(vcharts[index].chart)
        if record["categories"] != categories:
            raise VerifyError("the categories read back differ from the data")
        if [s["name"] for s in record["series"]] != expected_names:
            raise VerifyError("the series names read back differ from the data")
        if [s["values"] for s in record["series"]] != expected_values:
            raise VerifyError("the series values read back differ from the data")

    commit_deck(prs, out, verify)
    result = {"output": str(out), "slide": number, "series": len(series),
              "categories": len(categories), "verified": True}
    emit_action(args, result, [f"Wrote {out}",
                               f"Updated chart on slide {number}: {len(categories)} categories, {len(series)} series."])
    return 0


# ---------------------------------------------------------------- check-text

def normalise(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def read_phrases(path: Path) -> list:
    if not path.is_file():
        raise ToolError(f"Input file not found: {path}", 1)
    try:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    except UnicodeDecodeError as err:
        raise ToolError(f"Could not read {path}: {err}", 1)
    return [p for p in (normalise(line) for line in lines) if p]


def render_check_text(result: dict) -> None:
    for item in result["found"]:
        _print(f'Found "{item["phrase"]}" on slide {item["slide"]}')
    for phrase in result["missing"]:
        _print(f'Missing "{phrase}"')


def cmd_check_text(args) -> int:
    prs = open_deck(Path(args.deck))
    phrases = read_phrases(Path(args.expect))
    slide_texts = []
    for slide in prs.slides:
        texts, tables, charts = slide_contents(slide)
        pieces = list(texts)
        for table in tables:
            for row in table:
                pieces.extend(row)
        for chart in charts:
            pieces.extend(chart["categories"])
            pieces.extend(s["name"] for s in chart["series"])
        if args.include_notes:
            pieces.append(read_notes(slide))
        slide_texts.append(normalise(" ".join(pieces)))
    found, missing = [], []
    for phrase in phrases:
        number = next((i for i, text in enumerate(slide_texts, start=1) if phrase in text), None)
        if number is None:
            missing.append(phrase)
        else:
            found.append({"phrase": phrase, "slide": number})
    result = {"found": found, "missing": missing}
    emit_read(args, result, render_check_text)
    if missing:
        raise ToolError(f"{len(missing)} phrase(s) missing: {', '.join(missing)}", 1)
    return 0


# ---------------------------------------------------------------- validate

def _rels_name(part: str) -> str:
    directory, base = posixpath.split(part)
    return posixpath.join(directory, "_rels", base + ".rels")


def _rels_source(rels: str) -> str:
    """The part a .rels file belongs to; '' for the package root."""
    directory, base = posixpath.split(rels)
    owner_dir = posixpath.dirname(directory)
    name = base[: -len(".rels")]
    if not name:
        return ""
    return posixpath.join(owner_dir, name) if owner_dir else name


def _resolve(source: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    target = target.split("#", 1)[0]
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def _rel_entries(docs: dict, part: str) -> list:
    root = docs.get(_rels_name(part))
    if root is None:
        return []
    entries = []
    for rel in root.iter(f"{{{NS_PKG}}}Relationship"):
        entries.append({
            "id": rel.get("Id"),
            "type": rel.get("Type", ""),
            "target": rel.get("Target", ""),
            "external": rel.get("TargetMode") == "External",
        })
    return entries


def run_validate_checks(path: Path) -> list:
    checks: list = []

    def record(name: str, ok: bool, detail: str = "") -> bool:
        checks.append({"name": name, "status": "pass" if ok else "fail", "detail": detail})
        return ok

    def skip(names: list) -> None:
        for name in names:
            checks.append({"name": name, "status": "skipped", "detail": "not run because an earlier check failed"})

    names_in_order = VALIDATE_CHECKS
    try:
        zf = zipfile.ZipFile(path)
    except Exception as err:
        record(names_in_order[0], False, f"cannot open as a zip: {type(err).__name__}")
        skip(names_in_order[1:])
        return checks
    with zf:
        try:
            bad_member = zf.testzip()
        except Exception as err:
            bad_member = type(err).__name__
        if not record(names_in_order[0], bad_member is None,
                      "" if bad_member is None else f"corrupt member: {bad_member}"):
            skip(names_in_order[1:])
            return checks

        names = set(zf.namelist())
        ok_ct = "[Content_Types].xml" in names and "ppt/presentation.xml" in names
        record(names_in_order[1], ok_ct,
               "" if ok_ct else "missing [Content_Types].xml or ppt/presentation.xml")

        docs: dict = {}
        bad_xml = []
        for name in sorted(names):
            if name.endswith((".xml", ".rels")):
                try:
                    docs[name] = etree.fromstring(zf.read(name), XML_PARSER)
                except Exception:
                    bad_xml.append(name)
        record(names_in_order[2], not bad_xml, ("unparsable: " + ", ".join(bad_xml)) if bad_xml else "")

        missing_targets = []
        for name, root in docs.items():
            if not name.endswith(".rels"):
                continue
            source = _rels_source(name)
            for rel in root.iter(f"{{{NS_PKG}}}Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = _resolve(source, rel.get("Target", ""))
                if target not in names:
                    missing_targets.append(f"{name} -> {rel.get('Target')}")
        record(names_in_order[3], not missing_targets, "; ".join(missing_targets))

        try:
            Presentation(str(path))
            opened, detail = True, ""
        except Exception as err:
            opened, detail = False, f"{type(err).__name__}: {err}"
        record(names_in_order[4], opened, detail)

        pres = docs.get("ppt/presentation.xml")
        if not ok_ct or pres is None:
            skip(names_in_order[5:])
            return checks

        sld_ids = list(pres.iter(f"{{{NS_P}}}sldId"))
        try:
            ids = [int(e.get("id")) for e in sld_ids]
            ids_ok = len(set(ids)) == len(ids) and all(i >= 256 for i in ids)
            id_detail = "" if ids_ok else f"ids: {ids}"
        except (TypeError, ValueError):
            ids_ok, id_detail = False, "a slide id is not a number"
        record(names_in_order[5], ids_ok, id_detail)

        pres_rels = {e["id"]: e for e in _rel_entries(docs, "ppt/presentation.xml")}
        slide_parts = []
        for e in sld_ids:
            entry = pres_rels.get(e.get(f"{{{NS_R}}}id"))
            if entry is not None and entry["type"].endswith("/slide"):
                slide_parts.append(_resolve("ppt/presentation.xml", entry["target"]))

        no_layout = [s for s in slide_parts
                     if not any(r["type"].endswith("/slideLayout") for r in _rel_entries(docs, s))]
        record(names_in_order[6], not no_layout, "no layout relationship: " + ", ".join(no_layout) if no_layout else "")

        broken_notes = []
        for slide in slide_parts:
            for rel in _rel_entries(docs, slide):
                if not rel["type"].endswith("/notesSlide"):
                    continue
                notes_part = _resolve(slide, rel["target"])
                back = [r for r in _rel_entries(docs, notes_part)
                        if r["type"].endswith("/slide") and _resolve(notes_part, r["target"]) == slide]
                if not back:
                    broken_notes.append(notes_part)
        record(names_in_order[7], not broken_notes,
               "notes do not point back: " + ", ".join(broken_notes) if broken_notes else "")

        dangling = []
        for slide in slide_parts:
            root = docs.get(slide)
            if root is None:
                dangling.append(f"{slide} unreadable")
                continue
            rel_ids = {r["id"] for r in _rel_entries(docs, slide)}
            for blip in root.iter(f"{{{NS_A}}}blip"):
                for attr in (f"{{{NS_R}}}embed", f"{{{NS_R}}}link"):
                    value = blip.get(attr)
                    if value and value not in rel_ids:
                        dangling.append(f"{slide} -> {value}")
        record(names_in_order[8], not dangling, "; ".join(dangling))
    return checks


def cmd_validate(args) -> int:
    path = Path(args.deck)
    if not path.is_file():
        raise ToolError(f"Input file not found: {path}", 1)
    check_format(path)
    checks = run_validate_checks(path)
    failed = [c for c in checks if c["status"] == "fail"]
    result = {"valid": not failed, "checks": checks}

    def render(data: dict) -> None:
        for check in data["checks"]:
            label = check["status"].upper()
            suffix = f": {check['detail']}" if check["detail"] else ""
            _print(f"{label} {check['name']}{suffix}")
        _print("Verdict: " + ("valid" if data["valid"] else "invalid"))

    emit_read(args, result, render)
    if failed:
        raise ToolError(f"{len(failed)} check(s) failed. See checks.", 1)
    return 0


# ---------------------------------------------------------------- refused features

def cmd_unsupported(args) -> int:
    feature, alternative = UNSUPPORTED_COMMANDS[args.command]
    raise ToolError(f"Not supported: {feature}. {alternative}", 2)


# ---------------------------------------------------------------- argument parsing

class _Parser(argparse.ArgumentParser):
    """Argument errors are exit 1 and name the option, as the spec requires."""

    def error(self, message):
        _eprint(message)
        raise SystemExit(1)


def _add_common(parser, force=True, json_flag=True, out=True):
    if out:
        parser.add_argument("--out", help="write the result to this new file")
    if force:
        parser.add_argument("--force", action="store_true", help="allow replacing an existing output file")
    if json_flag:
        parser.add_argument("--json", action="store_true", help="print the result as JSON on stdout")


def _add_change(parser):
    parser.add_argument("--out", help="output deck (default: <name>.<command>.pptx next to the input)")
    parser.add_argument("--in-place", action="store_true", help="change the input file itself")
    parser.add_argument("--force", action="store_true", help="allow replacing an existing output file")
    parser.add_argument("--json", action="store_true", help="print the result as JSON on stdout")


def build_parser():
    parser = _Parser(prog="pptx_tool.py", description="Create, read and change PowerPoint decks.",
                     allow_abbrev=False)
    parser.add_argument("--version", action="version", version=VERSION)
    sub = parser.add_subparsers(dest="command", metavar="COMMAND", required=True, parser_class=_Parser)

    p = sub.add_parser("create", help="make a deck from a text outline")
    p.add_argument("outline", help="UTF-8 outline .txt")
    p.add_argument("--out", required=True, help="new deck, must end in .pptx")
    p.add_argument("--template", help="optional .pptx whose masters and layouts are used")
    p.add_argument("--force", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_create)

    p = sub.add_parser("read", help="print slide text, tables, charts and notes")
    p.add_argument("deck")
    p.add_argument("--slides", help="slide list such as 1-3,5,8- (default all)")
    _add_common(p)
    p.set_defaults(func=cmd_read)

    p = sub.add_parser("replace", help="replace text in slides, tables and notes")
    p.add_argument("deck")
    p.add_argument("--find", help="text to find (case-sensitive)")
    p.add_argument("--replace", help="replacement text")
    p.add_argument("--pairs", help="JSON object mapping find text to replacement text")
    p.add_argument("--include-notes", action=argparse.BooleanOptionalAction, default=True)
    _add_change(p)
    p.set_defaults(func=cmd_replace)

    p = sub.add_parser("add", help="add a slide")
    p.add_argument("deck")
    p.add_argument("--at", type=int, help="insert before this slide number (default: at end)")
    p.add_argument("--layout", default=TITLE_LAYOUT)
    p.add_argument("--title")
    p.add_argument("--body", action="append", help="one bullet; '- ' level 1, two spaces + '- ' level 2 (repeatable)")
    p.add_argument("--notes")
    _add_change(p)
    p.set_defaults(func=cmd_add)

    p = sub.add_parser("delete", help="delete slides")
    p.add_argument("deck")
    p.add_argument("--slides", required=True, help="slide list such as 2 or 4-6")
    _add_change(p)
    p.set_defaults(func=cmd_delete)

    p = sub.add_parser("reorder", help="change slide order")
    p.add_argument("deck")
    p.add_argument("--order", required=True, help="every slide number once, e.g. 3,1,2")
    _add_change(p)
    p.set_defaults(func=cmd_reorder)

    p = sub.add_parser("duplicate", help="copy a slide to the position after it")
    p.add_argument("deck")
    p.add_argument("--slide", type=int, required=True)
    _add_change(p)
    p.set_defaults(func=cmd_duplicate)

    p = sub.add_parser("split", help="split a deck into parts")
    p.add_argument("deck")
    group = p.add_mutually_exclusive_group()
    group.add_argument("--every", type=int, help="slides per part")
    group.add_argument("--ranges", help='ranges such as "1-3,4-6"')
    p.add_argument("--out-dir", help="folder for the parts (must be new or empty)")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_split)

    p = sub.add_parser("replace-image", help="swap a picture on a slide")
    p.add_argument("deck")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", help="picture shape name")
    p.add_argument("--shape-index", type=int, help="picture number on the slide, counted from 1 in z-order")
    p.add_argument("--image", required=True, help="PNG, JPEG, GIF, BMP or TIFF file")
    _add_change(p)
    p.set_defaults(func=cmd_replace_image)

    p = sub.add_parser("chart-data", help="replace data of a category chart")
    p.add_argument("deck")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", help="chart shape name (needed when the slide has several charts)")
    p.add_argument("--data", required=True, help="JSON file with categories and series")
    _add_change(p)
    p.set_defaults(func=cmd_chart_data)

    p = sub.add_parser("check-text", help="check that phrases appear in the deck")
    p.add_argument("deck")
    p.add_argument("--expect", required=True, help="text file, one phrase per line")
    p.add_argument("--include-notes", action=argparse.BooleanOptionalAction, default=True)
    _add_common(p)
    p.set_defaults(func=cmd_check_text)

    p = sub.add_parser("validate", help="check the package structure of a deck")
    p.add_argument("deck")
    _add_common(p)
    p.set_defaults(func=cmd_validate)

    # Refused commands: no help text, so they are not listed in --help; they accept any arguments and exit 2.
    for name in UNSUPPORTED_COMMANDS:
        p = sub.add_parser(name)
        p.add_argument("rest", nargs=argparse.REMAINDER)
        p.set_defaults(func=cmd_unsupported)
    return parser


def main(argv=None) -> int:
    _configure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args) or 0)
    except ToolError as err:
        _eprint(err.message)
        return err.code
    except PermissionError:
        _eprint(LOCKED_MSG)
        return 1
    except Exception as err:  # any unexpected failure
        if os.environ.get("PPTX_TOOL_DEBUG"):
            traceback.print_exc()
        _eprint(f"Internal error ({type(err).__name__}). No output was kept.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
