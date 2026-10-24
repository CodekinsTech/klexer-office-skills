"""Acceptance tests for pptx_tool.py (spec section 3) plus checks for the open items in spec section 7.

Decks are built inside the tests with python-pptx. The tool is run as a
subprocess so the exit codes and printed messages are tested as callers see them.
"""
import hashlib
import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from PIL import Image as PILImage
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu, Inches

TOOL = Path(__file__).resolve().parents[1] / "scripts" / "pptx_tool.py"


def run(*args):
    return subprocess.run(
        [sys.executable, str(TOOL), *[str(a) for a in args]],
        capture_output=True, text=True, encoding="utf-8",
    )


def run_json(*args):
    result = run(*args, "--json")
    assert result.returncode == 0, result.stderr
    return json.loads(result.stdout)


def write_outline(tmp_path, text, name="outline.txt"):
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return path


def make_titled_deck(path, titles):
    prs = Presentation()
    layout = prs.slide_layouts[1]
    for title in titles:
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text_frame.text = title
    prs.save(str(path))
    return path


def slide_titles(path):
    prs = Presentation(str(path))
    return [s.shapes.title.text_frame.text if s.shapes.title is not None else "" for s in prs.slides]


def all_zip_text(path):
    chunks = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if name.endswith(".xml"):
                chunks.append(zf.read(name).decode("utf-8", "replace"))
    return "\n".join(chunks)


CREATE_OUTLINE = (
    "# Intro\n"
    "- one\n"
    "  - sub\n"
    "Notes: hello\n"
    "\n"
    "# End\n"
    "- bye\n"
)


def build_create_deck(tmp_path):
    outline = write_outline(tmp_path, CREATE_OUTLINE)
    out = tmp_path / "t_create.pptx"
    assert run("create", outline, "--out", out).returncode == 0
    return out


# ------------------------------------------------------------------ create / read

def test_create(tmp_path):
    outline = write_outline(tmp_path, CREATE_OUTLINE)
    out = tmp_path / "t_create.pptx"
    result = run("create", outline, "--out", out)
    assert result.returncode == 0, result.stderr
    prs = Presentation(str(out))
    assert len(prs.slides) == 2
    first, second = prs.slides
    assert first.shapes.title.text_frame.text == "Intro"
    body = [(p.level, p.text) for p in first.placeholders[1].text_frame.paragraphs]
    assert body == [(0, "one"), (1, "sub")]
    assert first.notes_slide.notes_text_frame.text == "hello"
    assert second.shapes.title.text_frame.text == "End"
    assert second.placeholders[1].text_frame.text == "bye"


def test_create_text_before_title(tmp_path):
    outline = write_outline(tmp_path, "- x\n# Later\n")
    result = run("create", outline, "--out", tmp_path / "bad.pptx")
    assert result.returncode == 1
    assert 'Line 1 of the outline has text before the first "# " title.' in result.stderr
    assert not (tmp_path / "bad.pptx").exists()


def test_read(tmp_path):
    out = build_create_deck(tmp_path)
    data = run_json("read", out)
    assert data["slide_count"] == 2
    assert data["slides"][0]["title"] == "Intro"
    assert data["slides"][0]["notes"] == "hello"
    assert data["slides"][1]["notes"] == ""


def test_read_chart(tmp_path):
    path = tmp_path / "chart.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    data = CategoryChartData()
    data.categories = ["Q1", "Q2"]
    data.add_series("Sales", (10, 12))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(6), Inches(4), data)
    prs.save(str(path))
    result = run_json("read", path)
    chart = result["slides"][0]["charts"][0]
    assert chart["categories"] == ["Q1", "Q2"]
    assert chart["series"] == [{"name": "Sales", "values": [10, 12]}]


# ------------------------------------------------------------------ replace

def build_replace_deck(path):
    prs = Presentation()
    layout = prs.slide_layouts[5]
    s1 = prs.slides.add_slide(layout)
    s1.shapes.title.text_frame.text = "ACME Corp"
    table = s1.shapes.add_table(1, 1, Inches(1), Inches(2), Inches(4), Inches(1)).table
    table.cell(0, 0).text = "ACME"
    s2 = prs.slides.add_slide(layout)
    box = s2.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    para = box.text_frame.paragraphs[0]
    first = para.add_run()
    first.text = "AC"
    first.font.bold = True
    second = para.add_run()
    second.text = "ME"
    s2.notes_slide.notes_text_frame.text = "ACME notes"
    prs.save(str(path))
    return path


def test_replace(tmp_path):
    src = build_replace_deck(tmp_path / "rep.pptx")
    result = run_json("replace", src, "--find", "ACME", "--replace", "Globex")
    assert result["replacements"] == 4
    assert result["notes_replacements"] == 1
    assert result["per_slide"] == {"1": 2, "2": 2}
    assert "speaker notes" not in result["not_searched"]
    out = Path(result["output"])
    assert out == tmp_path / "rep.replace.pptx"
    text = all_zip_text(out)
    assert "ACME" not in text
    assert text.count("Globex") == 4
    # the split match kept the formatting of its first run
    prs = Presentation(str(out))
    run_el = prs.slides[1].shapes[-1].text_frame.paragraphs[0].runs[0]
    assert run_el.text == "Globex" and run_el.font.bold is True
    assert src.exists() and "ACME" in all_zip_text(src)


def test_replace_notes_excluded(tmp_path):
    src = build_replace_deck(tmp_path / "rep.pptx")
    result = run_json("replace", src, "--find", "ACME", "--replace", "Globex", "--no-include-notes")
    assert result["replacements"] == 3
    assert result["notes_replacements"] == 0
    assert "speaker notes" in result["not_searched"]


def test_replace_zero_hits_still_writes(tmp_path):
    src = build_replace_deck(tmp_path / "rep.pptx")
    result = run_json("replace", src, "--find", "NOPE", "--replace", "X")
    assert result["replacements"] == 0
    assert Path(result["output"]).exists()


# ------------------------------------------------------------------ add

def test_add(tmp_path):
    out = build_create_deck(tmp_path)
    added = tmp_path / "added.pptx"
    result = run("add", out, "--at", "2", "--title", "New", "--body", "- point", "--out", added)
    assert result.returncode == 0, result.stderr
    prs = Presentation(str(added))
    assert len(prs.slides) == 3
    assert prs.slides[1].shapes.title.text_frame.text == "New"
    assert prs.slides[1].placeholders[1].text_frame.text == "point"


def test_add_position_out_of_range(tmp_path):
    out = build_create_deck(tmp_path)
    result = run("add", out, "--at", "9", "--title", "X", "--out", tmp_path / "x.pptx")
    assert result.returncode == 1
    assert "Position must be between 1 and 3" in result.stderr


# ------------------------------------------------------------------ delete

def test_delete(tmp_path):
    src = make_titled_deck(tmp_path / "three.pptx", ["A", "B", "C"])
    result = run_json("delete", src, "--slides", "2")
    assert Path(result["output"]).exists()
    assert slide_titles(result["output"]) == ["A", "C"]
    # the deleted slide's text is not left in the package
    assert "<a:t>B</a:t>" not in all_zip_text(result["output"])


def test_delete_all_slides_refused(tmp_path):
    src = make_titled_deck(tmp_path / "three.pptx", ["A", "B", "C"])
    result = run("delete", src, "--slides", "1-3", "--out", tmp_path / "x.pptx")
    assert result.returncode == 1
    assert "A deck must keep at least one slide." in result.stderr


def test_delete_then_add_keeps_part_names_unique(tmp_path):
    src = make_titled_deck(tmp_path / "three.pptx", ["A", "B", "C"])
    deleted = Path(run_json("delete", src, "--slides", "1")["output"])
    added = tmp_path / "again.pptx"
    assert run("add", deleted, "--title", "D", "--out", added).returncode == 0
    assert slide_titles(added) == ["B", "C", "D"]


# ------------------------------------------------------------------ reorder

def test_reorder(tmp_path):
    src = make_titled_deck(tmp_path / "abc.pptx", ["A", "B", "C"])
    result = run_json("reorder", src, "--order", "3,1,2")
    assert slide_titles(result["output"]) == ["C", "A", "B"]


def test_reorder_not_a_permutation(tmp_path):
    src = make_titled_deck(tmp_path / "abc.pptx", ["A", "B", "C"])
    result = run("reorder", src, "--order", "1,1,2", "--out", tmp_path / "x.pptx")
    assert result.returncode == 1
    assert "Order must list every slide 1..3 exactly once." in result.stderr


# ------------------------------------------------------------------ duplicate

def test_duplicate(tmp_path):
    outline = write_outline(tmp_path, "# First\n- a\n- b\n# Second\n")
    src = tmp_path / "two.pptx"
    assert run("create", outline, "--out", src).returncode == 0
    result = run_json("duplicate", src, "--slide", "1")
    dup = Presentation(result["output"])
    assert len(dup.slides) == 3
    assert dup.slides[1].shapes.title.text_frame.text == "First"
    assert dup.slides[1].placeholders[1].text_frame.text == "a\nb"


def test_duplicate_chart_refused(tmp_path):
    src = tmp_path / "chart.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    data = CategoryChartData()
    data.categories = ["Q1"]
    data.add_series("S", (1,))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(6), Inches(4), data)
    prs.save(str(src))
    result = run("duplicate", src, "--slide", "1")
    assert result.returncode == 2
    assert "chart" in result.stderr


# ------------------------------------------------------------------ split

def test_split(tmp_path):
    src = make_titled_deck(tmp_path / "five.pptx", ["S1", "S2", "S3", "S4", "S5"])
    out_dir = tmp_path / "parts"
    result = run_json("split", src, "--every", "2", "--out-dir", out_dir)
    assert len(result["parts"]) == 3
    counts = [len(Presentation(p["file"]).slides) for p in result["parts"]]
    assert counts == [2, 2, 1]
    assert slide_titles(result["parts"][1]["file"]) == ["S3", "S4"]


def test_split_refuses_non_empty_folder(tmp_path):
    src = make_titled_deck(tmp_path / "five.pptx", ["S1", "S2"])
    out_dir = tmp_path / "parts"
    out_dir.mkdir()
    (out_dir / "keep.txt").write_text("x")
    result = run("split", src, "--every", "1", "--out-dir", out_dir)
    assert result.returncode == 1
    assert f"Output folder is not empty: {out_dir}" in result.stderr


# ------------------------------------------------------------------ replace-image

def _png(path, colour):
    PILImage.new("RGB", (100, 100), colour).save(path, "PNG")
    return path


def test_replace_image(tmp_path):
    red = _png(tmp_path / "red.png", (255, 0, 0))
    blue = _png(tmp_path / "blue.png", (0, 0, 255))
    src = tmp_path / "pic.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    old = slide.shapes.add_picture(str(red), Inches(1), Inches(1), Inches(2), Inches(2))
    box = (old.left, old.top, old.width, old.height)
    prs.save(str(src))
    result = run_json("replace-image", src, "--slide", "1", "--image", blue)
    out = Presentation(result["output"])
    pictures = [s for s in out.slides[0].shapes if s.shape_type == 13]
    assert len(pictures) == 1
    assert hashlib.sha256(pictures[0].image.blob).hexdigest() == hashlib.sha256(blue.read_bytes()).hexdigest()
    assert (pictures[0].left, pictures[0].top, pictures[0].width, pictures[0].height) == box


def test_replace_image_bad_format(tmp_path):
    src = tmp_path / "pic.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(_png(tmp_path / "r.png", (1, 2, 3))), 0, 0, Inches(1), Inches(1))
    prs.save(str(src))
    text_file = tmp_path / "not_image.txt"
    text_file.write_text("hello")
    result = run("replace-image", src, "--slide", "1", "--image", text_file)
    assert result.returncode == 1
    assert f"Image must be PNG, JPEG, GIF, BMP or TIFF: {text_file}" in result.stderr


# ------------------------------------------------------------------ chart-data

def build_column_chart(path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    data = CategoryChartData()
    data.categories = ["Q1", "Q2"]
    data.add_series("Sales", (10, 12))
    frame = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(6), Inches(4), data)
    frame.chart.plots[0].series[0].format.fill.solid()
    frame.chart.plots[0].series[0].format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
    prs.save(str(path))
    return path


def test_chart_data(tmp_path):
    src = build_column_chart(tmp_path / "chart.pptx")
    data = tmp_path / "new.json"
    data.write_text(json.dumps({"categories": ["Q1", "Q2", "Q3"],
                                "series": [{"name": "Sales", "values": [5, 6, 7]}]}), encoding="utf-8")
    result = run_json("chart-data", src, "--slide", "1", "--data", data)
    chart = run_json("read", result["output"])["slides"][0]["charts"][0]
    assert chart["categories"] == ["Q1", "Q2", "Q3"]
    assert chart["series"] == [{"name": "Sales", "values": [5, 6, 7]}]
    # per-series colour survives the data replacement (spec open item)
    prs = Presentation(result["output"])
    chart_obj = [s for s in prs.slides[0].shapes if s.has_chart][0].chart
    assert str(chart_obj.plots[0].series[0].format.fill.fore_color.rgb) == "FF0000"


def test_chart_data_count_mismatch(tmp_path):
    src = build_column_chart(tmp_path / "chart.pptx")
    data = tmp_path / "bad.json"
    data.write_text(json.dumps({"categories": ["Q1", "Q2"],
                                "series": [{"name": "Sales", "values": [1, 2, 3]}]}), encoding="utf-8")
    result = run("chart-data", src, "--slide", "1", "--data", data, "--out", tmp_path / "x.pptx")
    assert result.returncode == 1
    assert 'Series "Sales" has 3 values for 2 categories.' in result.stderr


# ------------------------------------------------------------------ check-text

def test_check_text(tmp_path):
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[5])
    s1.shapes.title.text_frame.text = "Revenue"
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.notes_slide.notes_text_frame.text = "Summary"
    src = tmp_path / "check.pptx"
    prs.save(str(src))
    expect = tmp_path / "expect.txt"
    expect.write_text("Revenue\nSummary\n\nForecast\n", encoding="utf-8")
    result = run("check-text", src, "--expect", expect, "--json")
    assert result.returncode == 1
    data = json.loads(result.stdout)
    assert data["found"] == [{"phrase": "Revenue", "slide": 1}, {"phrase": "Summary", "slide": 2}]
    assert data["missing"] == ["Forecast"]


# ------------------------------------------------------------------ validate

def test_validate_valid(tmp_path):
    src = build_create_deck(tmp_path)
    result = run_json("validate", src)
    assert result["valid"] is True


def test_validate_truncated(tmp_path):
    src = build_create_deck(tmp_path)
    cut = tmp_path / "cut.pptx"
    cut.write_bytes(src.read_bytes()[:1000])
    result = run("validate", cut, "--json")
    assert result.returncode == 1
    assert json.loads(result.stdout)["valid"] is False


# ------------------------------------------------------------------ general rules

def test_missing_option_names_option(tmp_path):
    src = make_titled_deck(tmp_path / "a.pptx", ["A"])
    result = run("delete", src)
    assert result.returncode == 1
    assert "--slides" in result.stderr


def test_existing_output_needs_force(tmp_path):
    src = make_titled_deck(tmp_path / "a.pptx", ["A", "B"])
    target = tmp_path / "taken.pptx"
    target.write_bytes(b"occupied")
    result = run("delete", src, "--slides", "1", "--out", target)
    assert result.returncode == 1
    assert f"Output already exists: {target}. Choose another --out or pass --force." in result.stderr
    assert target.read_bytes() == b"occupied"
    assert run("delete", src, "--slides", "1", "--out", target, "--force").returncode == 0


def test_macro_enabled_refused(tmp_path):
    src = make_titled_deck(tmp_path / "a.pptx", ["A"])
    macro = tmp_path / "a.pptm"
    macro.write_bytes(src.read_bytes())
    result = run("read", macro)
    assert result.returncode == 2
    assert "Macro-enabled presentations are not supported." in result.stderr


def test_legacy_ppt_refused(tmp_path):
    legacy = tmp_path / "old.ppt"
    legacy.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 100)
    result = run("read", legacy)
    assert result.returncode == 2
    assert "Legacy .ppt is not supported. Save as .pptx and try again." in result.stderr


def test_encrypted_refused(tmp_path):
    enc = tmp_path / "locked.pptx"
    enc.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 50 + "EncryptedPackage".encode("utf-16-le"))
    result = run("read", enc)
    assert result.returncode == 2
    assert "This deck is password-protected." in result.stderr


def test_input_never_overwritten_without_in_place(tmp_path):
    src = make_titled_deck(tmp_path / "a.pptx", ["A", "B"])
    result = run("delete", src, "--slides", "1", "--out", src)
    assert result.returncode == 1
    assert "Refusing to overwrite the input. Use --in-place to change it." in result.stderr
    assert slide_titles(src) == ["A", "B"]


# ------------------------------------------------------------------ refused features

RENDER_MESSAGE = (
    "Not supported: rendering slides to PDF or images. "
    "Export the deck from PowerPoint with File > Export, or use a conversion tool outside this skill."
)


@pytest.mark.parametrize("command", ["render", "export", "pdf"])
def test_render_request_refused_with_exit_2(tmp_path, command):
    src = make_titled_deck(tmp_path / "deck.pptx", ["Render me"])
    target = tmp_path / "deck.pdf"
    result = run(command, src, "--format", "pdf", "--out", target)
    assert result.returncode == 2
    assert result.stdout == ""
    assert result.stderr.strip() == RENDER_MESSAGE
    assert not target.exists()
    assert slide_titles(src) == ["Render me"]


@pytest.mark.parametrize("command, feature", [
    ("animate", "animations"),
    ("transition", "transitions"),
    ("theme", "theme editing"),
    ("master", "slide master editing"),
])
def test_other_unsupported_features_refused_with_exit_2(tmp_path, command, feature):
    src = make_titled_deck(tmp_path / "deck.pptx", ["A"])
    result = run(command, src)
    assert result.returncode == 2
    assert result.stderr.startswith(f"Not supported: {feature}. ")


def test_unknown_command_is_usage_error_exit_1(tmp_path):
    src = make_titled_deck(tmp_path / "deck.pptx", ["A"])
    result = run("frobnicate", src)
    assert result.returncode == 1
    assert "invalid choice: 'frobnicate'" in result.stderr
    assert "Not supported" not in result.stderr


def test_misspelled_command_is_usage_error_exit_1(tmp_path):
    src = make_titled_deck(tmp_path / "deck.pptx", ["A"])
    result = run("rendr", src)
    assert result.returncode == 1
    assert "Not supported" not in result.stderr
