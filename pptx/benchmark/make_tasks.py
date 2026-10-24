#!/usr/bin/env python3
"""Build the PPTX v2 benchmark cases.

Writes one folder per task under cases/. Each folder holds the task's input
files (decks built with python-pptx, outlines, JSON data, expect files, PNG
images) and task.json, the task definition: the user request, the CLI
arguments and the expected result.

The expected results are written by hand from specs/pptx.md and SKILL.md.
They are never produced by running the tool. run_benchmark.py only compares
what the tool does with these expectations.

Usage (use the venv that has python-pptx):
    python make_tasks.py
"""
import hashlib
import json
import shutil
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

HERE = Path(__file__).resolve().parent
CASES = HERE / "cases"

# OLE2 compound-file signature with no content. A legacy .ppt starts like this.
OLE_HEADER = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504

CASE_LIST = []


def case(case_id, category, title, request):
    def register(fn):
        CASE_LIST.append({"id": case_id, "category": category,
                          "title": title, "request": request, "build": fn})
        return fn
    return register


# ---------------------------------------------------------------- helpers

def _layout(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise KeyError(name)


def _set_name(shape, name):
    shape._element.xpath("./*[1]/p:cNvPr")[0].set("name", name)


def add_slide(prs, title, bullets=(), notes=None, layout="Title and Content"):
    """bullets: list of (text, level) with level 1 or 2."""
    slide = prs.slides.add_slide(_layout(prs, layout))
    if title is not None:
        slide.shapes.title.text = title
    if bullets:
        tf = slide.placeholders[1].text_frame
        for i, (text, level) in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = text
            p.level = level - 1
    if notes is not None:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


def add_table(slide, rows):
    shape = slide.shapes.add_table(len(rows), len(rows[0]),
                                   Inches(1), Inches(3), Inches(6), Inches(1.5))
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            shape.table.cell(r, c).text = text
    return shape


def add_column_chart(slide, categories, series, name):
    cd = CategoryChartData()
    cd.categories = categories
    for series_name, values in series:
        cd.add_series(series_name, values)
    frame = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                   Inches(1), Inches(2.5), Inches(6), Inches(4), cd)
    _set_name(frame, name)
    return frame


def make_png(path, rgb):
    Image.new("RGB", (100, 100), rgb).save(path, format="PNG")


def sha256_file(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write_json(path, obj):
    Path(path).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def simple_deck(path, titles):
    prs = Presentation()
    for i, t in enumerate(titles, 1):
        add_slide(prs, t, bullets=[(f"body {i}", 1)])
    prs.save(path)


# ------------------------------------------------------------- create

@case("C1_create_outline_spec", "create",
      "Outline with levels, a speaker note and a second slide",
      "Make a deck from this outline: an intro slide with two bullets (one has a sub-bullet) "
      "and a speaker note, then an end slide with one bullet.")
def _c1(d):
    (d / "outline.txt").write_text(
        "# Intro\n- one\n  - sub\nNotes: hello\n\n# End\n- bye\n", encoding="utf-8")
    return (["create", "outline.txt", "--out", "t_create.pptx"],
            {"exit": 0,
             "outputs": {"t_create.pptx": {
                 "slide_count": 2,
                 "titles": ["Intro", "End"],
                 "body": {"1": [["one", 1], ["sub", 2]], "2": [["bye", 1]]},
                 "notes": {"1": "hello", "2": ""}}}})


@case("C2_create_text_before_title", "create",
      "Outline with a bullet before the first title is refused",
      "Make a deck from this outline. (The outline starts with a bullet, not a title.)")
def _c2(d):
    (d / "outline.txt").write_text("- x\n# Late\n- y\n", encoding="utf-8")
    return (["create", "outline.txt", "--out", "t_bad.pptx"],
            {"exit": 1,
             "stderr_has": ['Line 1 of the outline has text before the first "# " title.'],
             "outputs": {"t_bad.pptx": None}})


@case("C3_create_unicode_roundtrip", "create",
      "Thai, Arabic and accented text survives create",
      "Make a one-slide deck in Thai with a bullet in Arabic and one with accented Latin text.")
def _c3(d):
    text = ("# รายงานยอดขาย\n"
            "- ยอดขายไตรมาสที่ 1: 1,200 บาท\n"
            "- Café résumé naïve\n"
            "- مرحبا بالعالم\n"
            "Notes: หมายเหตุ ภาษาไทย\n")
    (d / "outline.txt").write_text(text, encoding="utf-8")
    return (["create", "outline.txt", "--out", "t_unicode.pptx"],
            {"exit": 0,
             "outputs": {"t_unicode.pptx": {
                 "slide_count": 1,
                 "titles": ["รายงานยอดขาย"],
                 "body": {"1": [["ยอดขายไตรมาสที่ 1: 1,200 บาท", 1],
                                ["Café résumé naïve", 1],
                                ["مرحبا بالعالم", 1]]},
                 "notes": {"1": "หมายเหตุ ภาษาไทย"}}}})


# --------------------------------------------------------------- read

@case("R1_read_column_chart", "read",
      "Read a column chart's categories, series and values",
      "What numbers are in the chart on my deck? Give me the JSON.")
def _r1(d):
    prs = Presentation()
    s = add_slide(prs, "Sales by quarter", layout="Title Only")
    add_column_chart(s, ["Q1", "Q2"], [("Sales", [10, 12])], "Sales chart")
    prs.save(d / "deck.pptx")
    return (["read", "deck.pptx", "--json"],
            {"exit": 0,
             "json": {"slide_count": 1,
                      "slides.0.title": "Sales by quarter",
                      "slides.0.charts.0.categories": ["Q1", "Q2"],
                      "slides.0.charts.0.series.0.name": "Sales",
                      "slides.0.charts.0.series.0.values": [10, 12],
                      "~slides.0.charts.0.type": "column"}})


@case("R2_read_table_and_slide_subset", "read",
      "Read one slide's table with --slides",
      "Show me only slide 2 of this deck, including its table, as JSON.")
def _r2(d):
    prs = Presentation()
    add_slide(prs, "Cover", layout="Title Only")
    s = add_slide(prs, "Regions", layout="Title Only")
    add_table(s, [["Region", "Revenue"], ["North", "120"], ["South", "95"]])
    add_slide(prs, "Close", layout="Title Only")
    prs.save(d / "deck.pptx")
    return (["read", "deck.pptx", "--slides", "2", "--json"],
            {"exit": 0,
             "json": {"slide_count": 3,
                      "len:slides": 1,
                      "slides.0.number": 2,
                      "slides.0.title": "Regions",
                      "slides.0.tables.0": [["Region", "Revenue"], ["North", "120"], ["South", "95"]]}})


@case("R3_read_never_creates_notes", "read",
      "Reading a deck does not change it or create a notes page",
      "Read this deck and tell me the notes on each slide. Do not change the file.")
def _r3(d):
    prs = Presentation()
    add_slide(prs, "Intro", notes="Speaker note A")
    add_slide(prs, "Body", bullets=[("x", 1)])
    prs.save(d / "deck.pptx")
    return (["read", "deck.pptx", "--json"],
            {"exit": 0,
             "json": {"slide_count": 2,
                      "slides.0.notes": "Speaker note A",
                      "slides.1.notes": "",
                      "slides.1.title": "Body"},
             "input_deck": {"file": "deck.pptx", "notes_pages": 1}})


# ------------------------------------------------------------ replace

@case("P1_replace_title_table_split_run_notes", "replace",
      "Replace across title, table cell, split-run text and notes",
      "Replace every ACME with Globex in this deck, including the notes, and tell me the counts.")
def _p1(d):
    prs = Presentation()
    s1 = add_slide(prs, "ACME Corp Overview", layout="Title Only")
    add_table(s1, [["Client", "ACME"], ["Region", "North"]])
    s2 = add_slide(prs, "Partners", notes="ACME notes for the partner call")
    para = s2.placeholders[1].text_frame.paragraphs[0]
    para.add_run().text = "Partner: AC"
    para.add_run().text = "ME Ltd"          # 'ACME' is split across two runs
    add_slide(prs, "acme lowercase stays", bullets=[("acme product line", 1)])
    prs.save(d / "deck.pptx")
    return (["replace", "deck.pptx", "--find", "ACME", "--replace", "Globex",
             "--out", "p1_out.pptx", "--json"],
            {"exit": 0,
             "json": {"replacements": 4, "notes_replacements": 1,
                      "per_slide": {"1": 2, "2": 2}, "verified": True},
             "outputs": {"p1_out.pptx": {
                 "slide_count": 3,
                 "titles": ["Globex Corp Overview", "Partners", "acme lowercase stays"],
                 "table": {"1": [["Client", "Globex"], ["Region", "North"]]},
                 "body": {"2": [["Partner: Globex Ltd", 1]],
                          "3": [["acme product line", 1]]},
                 "notes": {"2": "Globex notes for the partner call"},
                 "count_in_deck": {"Globex": 4, "ACME": 0},
                 "all_text_absent": ["ACME"]}}})


@case("P3_replace_pairs_file", "replace",
      "Several replacements from a pairs file, notes included",
      "Apply these renames to the deck (Q3 becomes Q4, Revenue becomes Income) and report the counts.")
def _p3(d):
    prs = Presentation()
    add_slide(prs, "Q3 Revenue review")
    add_slide(prs, "Summary", bullets=[("Revenue up in Q3", 1)], notes="Q3 notes")
    prs.save(d / "deck.pptx")
    write_json(d / "pairs.json", {"Q3": "Q4", "Revenue": "Income"})
    return (["replace", "deck.pptx", "--pairs", "pairs.json",
             "--out", "p3_out.pptx", "--json"],
            {"exit": 0,
             "json": {"replacements": 5, "notes_replacements": 1,
                      "per_slide": {"1": 2, "2": 3}, "verified": True},
             "outputs": {"p3_out.pptx": {
                 "slide_count": 2,
                 "titles": ["Q4 Income review", "Summary"],
                 "body": {"2": [["Income up in Q4", 1]]},
                 "notes": {"2": "Q4 notes"},
                 "all_text_absent": ["Q3", "Revenue"]}}})


@case("X7_replace_default_output_name", "replace",
      "Default output is a new file; the input is never overwritten",
      "Replace ACME with Globex in deck.pptx. I did not give an output name.")
def _x7(d):
    prs = Presentation()
    add_slide(prs, "ACME plan")
    prs.save(d / "deck.pptx")
    return (["replace", "deck.pptx", "--find", "ACME", "--replace", "Globex"],
            {"exit": 0,
             "outputs": {"deck.replace.pptx": {
                 "titles": ["Globex plan"],
                 "count_in_deck": {"Globex": 1, "ACME": 0}}},
             "input_deck": {"file": "deck.pptx", "titles": ["ACME plan"]}})


# ---------------------------------------------------------------- add

@case("A1_add_slide_levels_notes", "add",
      "Add a slide in the middle with sub-bullets and a note",
      "Add a 'Next steps' slide before slide 2 with three points, one of them a sub-point, "
      "and the note 'Discuss budget'.")
def _a1(d):
    prs = Presentation()
    add_slide(prs, "Alpha", bullets=[("a1", 1)])
    add_slide(prs, "Omega", bullets=[("o1", 1)])
    prs.save(d / "deck.pptx")
    return (["add", "deck.pptx", "--at", "2", "--title", "Next steps",
             "--body", "- Hire", "--body", "  - two engineers", "--body", "Plain line",
             "--notes", "Discuss budget", "--out", "a1_out.pptx", "--json"],
            {"exit": 0,
             "json": {"slide_count": 3, "position": 2, "verified": True},
             "outputs": {"a1_out.pptx": {
                 "slide_count": 3,
                 "titles": ["Alpha", "Next steps", "Omega"],
                 "body": {"1": [["a1", 1]],
                          "2": [["Hire", 1], ["two engineers", 2], ["Plain line", 1]],
                          "3": [["o1", 1]]},
                 "notes": {"2": "Discuss budget"}}}})


@case("A2_add_position_out_of_range", "add",
      "Add at a position past the end is refused",
      "Add a slide at position 5 in this three-slide deck.")
def _a2(d):
    simple_deck(d / "deck.pptx", ["One", "Two", "Three"])
    return (["add", "deck.pptx", "--at", "5", "--title", "X", "--out", "a2_out.pptx"],
            {"exit": 1, "stderr_nonempty": True, "outputs": {"a2_out.pptx": None}})


# ------------------------------------------------------------- delete

@case("D1_delete_slide_removes_part", "delete",
      "Delete slide 2 and check its content is gone from the file",
      "Delete slide 2 from this deck. Its text must not remain anywhere in the file.")
def _d1(d):
    prs = Presentation()
    add_slide(prs, "Keep one", bullets=[("first", 1)])
    add_slide(prs, "Drop ZQDROPZQ", bullets=[("ZQDROPZQ body", 1)], notes="ZQDROPZQ notes")
    add_slide(prs, "Keep three", bullets=[("third", 1)])
    prs.save(d / "deck.pptx")
    return (["delete", "deck.pptx", "--slides", "2", "--out", "d1_out.pptx"],
            {"exit": 0,
             "outputs": {"d1_out.pptx": {
                 "slide_count": 2,
                 "titles": ["Keep one", "Keep three"],
                 "package_absent": ["ZQDROPZQ"]}}})


# ------------------------------------------------------------ reorder

@case("O1_reorder_permutation", "reorder",
      "Move slide 3 to the front",
      "Reorder the slides so the order is 3, 1, 2.")
def _o1(d):
    prs = Presentation()
    for t in ["A", "B", "C"]:
        add_slide(prs, t, bullets=[(f"body {t}", 1)])
    prs.save(d / "deck.pptx")
    return (["reorder", "deck.pptx", "--order", "3,1,2", "--out", "o1_out.pptx", "--json"],
            {"exit": 0,
             "json": {"order": [3, 1, 2], "verified": True},
             "outputs": {"o1_out.pptx": {
                 "slide_count": 3,
                 "titles": ["C", "A", "B"],
                 "body": {"1": [["body C", 1]], "2": [["body A", 1]], "3": [["body B", 1]]}}}})


# ---------------------------------------------------------- duplicate

@case("U1_duplicate_text_notes_layout", "duplicate",
      "Duplicate slide 1 with bullets and notes",
      "Copy slide 1 and put the copy right after it.")
def _u1(d):
    prs = Presentation()
    add_slide(prs, "Plan", bullets=[("Goal", 1), ("Owner", 2)], notes="Plan notes")
    add_slide(prs, "Other", bullets=[("x", 1)])
    prs.save(d / "deck.pptx")
    return (["duplicate", "deck.pptx", "--slide", "1", "--out", "u1_out.pptx", "--json"],
            {"exit": 0,
             "json": {"source_slide": 1, "new_slide": 2, "verified": True},
             "outputs": {"u1_out.pptx": {
                 "slide_count": 3,
                 "titles": ["Plan", "Plan", "Other"],
                 "body": {"1": [["Goal", 1], ["Owner", 2]],
                          "2": [["Goal", 1], ["Owner", 2]],
                          "3": [["x", 1]]},
                 "notes": {"1": "Plan notes", "2": "Plan notes"}}}})


@case("U2_duplicate_chart_refused", "duplicate",
      "Duplicating a slide with a chart is refused and names the chart",
      "Copy slide 1 of this deck. Slide 1 has a chart.")
def _u2(d):
    prs = Presentation()
    s = add_slide(prs, "Chart slide", layout="Title Only")
    add_column_chart(s, ["Q1"], [("S", [1])], "c")
    add_slide(prs, "Plain", bullets=[("x", 1)])
    prs.save(d / "deck.pptx")
    return (["duplicate", "deck.pptx", "--slide", "1", "--out", "u2_out.pptx"],
            {"exit": 2,
             "stderr_has": ["Slide 1 contains a chart"],
             "outputs": {"u2_out.pptx": None}})


# --------------------------------------------------------------- split

@case("S1_split_every_two", "split",
      "Split a five-slide deck into parts of two",
      "Split this five-slide deck into files of two slides each.")
def _s1(d):
    simple_deck(d / "deck.pptx", ["S1", "S2", "S3", "S4", "S5"])
    return (["split", "deck.pptx", "--every", "2", "--out-dir", "parts", "--json"],
            {"exit": 0,
             "json": {"verified": True,
                      "len:parts": 3,
                      "parts.0.slides": 2,
                      "parts.1.first_slide": 3,
                      "parts.1.last_slide": 4,
                      "parts.2.slides": 1},
             "outputs": {
                 "parts/deck.part01.pptx": {"slide_count": 2, "titles": ["S1", "S2"]},
                 "parts/deck.part02.pptx": {"slide_count": 2, "titles": ["S3", "S4"]},
                 "parts/deck.part03.pptx": {"slide_count": 1, "titles": ["S5"]}},
             "dir_files": {"parts": 3}})


@case("S2_split_uncovered_slide", "split",
      "Split ranges that leave a slide out are refused",
      "Split this five-slide deck into slides 1-2 and 4-5 only.")
def _s2(d):
    simple_deck(d / "deck.pptx", ["S1", "S2", "S3", "S4", "S5"])
    return (["split", "deck.pptx", "--ranges", "1-2,4-5", "--out-dir", "parts_bad"],
            {"exit": 1,
             "stderr_has": ["Slide 3 is not covered"],
             "dir_files": {"parts_bad": 0}})


# ------------------------------------------------------- replace-image

@case("I1_replace_image_keeps_box_and_name", "replace-image",
      "Swap the one picture on a slide for a new image",
      "Replace the picture on slide 1 with blue.png.")
def _i1(d):
    make_png(d / "red.png", (255, 0, 0))
    make_png(d / "blue.png", (0, 0, 255))
    prs = Presentation()
    s = add_slide(prs, "Logo slide", layout="Title Only")
    pic = s.shapes.add_picture(str(d / "red.png"), Inches(1), Inches(2), Inches(3), Inches(3))
    _set_name(pic, "Logo")
    prs.save(d / "deck.pptx")
    return (["replace-image", "deck.pptx", "--slide", "1", "--image", "blue.png",
             "--out", "i1_out.pptx", "--json"],
            {"exit": 0,
             "json": {"slide": 1, "cropped_dropped": False, "verified": True},
             "outputs": {"i1_out.pptx": {
                 "slide_count": 1,
                 "pictures": {"1": [{"sha256": sha256_file(d / "blue.png"),
                                     "name": "Logo",
                                     "left": int(Inches(1)), "top": int(Inches(2)),
                                     "width": int(Inches(3)), "height": int(Inches(3))}]}}}})


# ---------------------------------------------------------- chart-data

@case("CD1_chart_data_update_keeps_type_and_colour", "chart-data",
      "Replace a column chart's data; type and series colour stay",
      "Update the chart on slide 1 with Q1 to Q3 and the Sales values 5, 6 and 7. Keep the colours.")
def _cd1(d):
    prs = Presentation()
    s = add_slide(prs, "Chart", layout="Title Only")
    gf = add_column_chart(s, ["Q1", "Q2"], [("Sales", [10, 12])], "Sales chart")
    ser = gf.chart.plots[0].series[0]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
    prs.save(d / "deck.pptx")
    write_json(d / "new_data.json",
               {"categories": ["Q1", "Q2", "Q3"], "series": [{"name": "Sales", "values": [5, 6, 7]}]})
    return (["chart-data", "deck.pptx", "--slide", "1", "--data", "new_data.json",
             "--out", "cd1_out.pptx", "--json"],
            {"exit": 0,
             "json": {"slide": 1, "verified": True},
             "outputs": {"cd1_out.pptx": {
                 "chart": {"1": {"type": "COLUMN_CLUSTERED",
                                 "categories": ["Q1", "Q2", "Q3"],
                                 "series": [["Sales", [5, 6, 7]]],
                                 "series_fill": ["FF0000"]}}}}})


@case("CD2_chart_data_count_mismatch", "chart-data",
      "Chart data with more values than categories is refused",
      "Put these values into the chart on slide 1: categories Q1 and Q2, Sales 5, 6 and 7.")
def _cd2(d):
    prs = Presentation()
    s = add_slide(prs, "Chart", layout="Title Only")
    add_column_chart(s, ["Q1", "Q2"], [("Sales", [10, 12])], "Sales chart")
    prs.save(d / "deck.pptx")
    write_json(d / "bad_data.json",
               {"categories": ["Q1", "Q2"], "series": [{"name": "Sales", "values": [5, 6, 7]}]})
    return (["chart-data", "deck.pptx", "--slide", "1", "--data", "bad_data.json",
             "--out", "cd2_out.pptx"],
            {"exit": 1,
             "stderr_has": ['Series "Sales" has 3 values for 2 categories.'],
             "outputs": {"cd2_out.pptx": None}})


@case("CD3_chart_data_xy_refused", "chart-data",
      "Editing an XY scatter chart is refused",
      "Change the data of the scatter chart on slide 1.")
def _cd3(d):
    prs = Presentation()
    s = add_slide(prs, "Scatter", layout="Title Only")
    cd = XyChartData()
    ser = cd.add_series("Points")
    ser.add_data_point(1.0, 2.0)
    ser.add_data_point(2.0, 3.0)
    s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, Inches(1), Inches(2), Inches(6), Inches(4), cd)
    prs.save(d / "deck.pptx")
    write_json(d / "data.json",
               {"categories": ["A", "B"], "series": [{"name": "Points", "values": [1, 2]}]})
    return (["chart-data", "deck.pptx", "--slide", "1", "--data", "data.json",
             "--out", "cd3_out.pptx"],
            {"exit": 2,
             "stderr_has": ["Only category charts (bar, column, line, pie) can be edited."],
             "outputs": {"cd3_out.pptx": None}})


# ---------------------------------------------------------- check-text

@case("K1_check_text_missing_phrase", "check-text",
      "Check phrases; one is missing and one is in the notes",
      "Check that this deck mentions Revenue, Summary and Forecast. Tell me what is missing.")
def _k1(d):
    prs = Presentation()
    add_slide(prs, "Revenue report")
    add_slide(prs, "Growth", bullets=[("Growth 4%", 1)], notes="Summary of plan")
    prs.save(d / "deck.pptx")
    (d / "expect.txt").write_text("Revenue\nSummary\nForecast\n\n", encoding="utf-8")
    return (["check-text", "deck.pptx", "--expect", "expect.txt", "--json"],
            {"exit": 1,
             "json": {"found": [{"phrase": "Revenue", "slide": 1},
                                {"phrase": "Summary", "slide": 2}],
                      "missing": ["Forecast"]}})


@case("K2_check_text_whitespace_and_chart", "check-text",
      "Whitespace is normalised; chart categories and series names count",
      "Check that this deck contains 'Net revenue growth', 'Alpha East' and 'Units sold'.")
def _k2(d):
    prs = Presentation()
    add_slide(prs, "Net   revenue\tgrowth", layout="Title Only")
    s2 = add_slide(prs, "Units", layout="Title Only")
    add_column_chart(s2, ["Alpha East", "Beta"], [("Units sold", [3, 4])], "Units chart")
    prs.save(d / "deck.pptx")
    (d / "expect.txt").write_text("Net revenue growth\nAlpha East\nUnits sold\n", encoding="utf-8")
    return (["check-text", "deck.pptx", "--expect", "expect.txt", "--json"],
            {"exit": 0,
             "json": {"found": [{"phrase": "Net revenue growth", "slide": 1},
                                {"phrase": "Alpha East", "slide": 2},
                                {"phrase": "Units sold", "slide": 2}],
                      "missing": []}})


# ------------------------------------------------------------ validate

@case("V1_validate_valid_deck", "validate",
      "A deck with notes and a picture passes every check",
      "Is this deck a valid file?")
def _v1(d):
    make_png(d / "red.png", (255, 0, 0))
    prs = Presentation()
    s = add_slide(prs, "Valid deck", bullets=[("point", 1)], notes="valid notes")
    s.shapes.add_picture(str(d / "red.png"), Inches(1), Inches(3), Inches(2), Inches(2))
    prs.save(d / "deck.pptx")
    return (["validate", "deck.pptx", "--json"],
            {"exit": 0,
             "json": {"valid": True, "len:checks": 9, "all:checks.status": "pass"}})


@case("V2_validate_truncated_zip", "validate",
      "A deck cut off at 1000 bytes fails validation",
      "Check whether broken.pptx is a valid deck. It was cut off during a download.")
def _v2(d):
    tmp = d / "_tmp.pptx"
    simple_deck(tmp, ["T1", "T2", "T3"])
    data = tmp.read_bytes()[:1000]
    tmp.unlink()
    (d / "broken.pptx").write_bytes(data)
    return (["validate", "broken.pptx", "--json"],
            {"exit": 1,
             "json": {"valid": False, "checks.0.status": "fail"}})


# ------------------------------------------------- refusal and safety

@case("X1_refusal_corrupt_file", "refusal",
      "A corrupt .pptx is refused cleanly",
      "Read corrupt.pptx and summarise it.")
def _x1(d):
    (d / "corrupt.pptx").write_bytes(b"PK\x03\x04" + b"this is not a real package " * 40)
    return (["read", "corrupt.pptx"],
            {"exit": 1,
             "stderr_has": ["File is not a readable .pptx (corrupt or truncated)"]})


@case("X2_refusal_missing_file", "refusal",
      "A missing input file is refused",
      "Read nope.pptx and list its slides.")
def _x2(d):
    return (["read", "nope.pptx"],
            {"exit": 1,
             "stderr_has": ["Input file not found", "nope.pptx"]})


@case("X3_refusal_legacy_ppt", "refusal",
      "A legacy .ppt is refused with exit 2",
      "Read old.ppt and list its slides.")
def _x3(d):
    (d / "old.ppt").write_bytes(OLE_HEADER)
    return (["read", "old.ppt"],
            {"exit": 2,
             "stderr_has": ["Legacy .ppt is not supported. Save as .pptx and try again."]})


@case("X4_refusal_macro_pptm", "refusal",
      "A macro-enabled .pptm is refused and not stripped",
      "Read macro.pptm and tell me the slide titles.")
def _x4(d):
    prs = Presentation()
    add_slide(prs, "Macro test")
    prs.save(d / "macro.pptm")
    return (["read", "macro.pptm"],
            {"exit": 2,
             "stderr_has": ["Macro-enabled presentations are not supported."]})


@case("X5_refusal_out_of_scope_render", "refusal",
      "A request to render slides to PDF is refused with exit 2",
      "Export deck.pptx to deck.pdf so I can email it.")
def _x5(d):
    simple_deck(d / "deck.pptx", ["Render me"])
    return (["render", "deck.pptx", "--format", "pdf", "--out", "deck.pdf"],
            {"exit": 2,
             "stderr_has": ["Not supported: rendering slides to PDF or images. "],
             "outputs": {"deck.pdf": None},
             "note": "Decision: a request for a feature the tool does not provide is a refusal, exit 2, "
                     "printing 'Not supported: <feature>. <alternative>.' An unknown command name or a "
                     "usage typo is exit 1. 'render' is a refused feature name, so exit 2."})


# -------------------------------------------------------------- build

def main():
    if CASES.exists():
        shutil.rmtree(CASES)
    CASES.mkdir(parents=True)
    for spec in CASE_LIST:
        folder = CASES / spec["id"]
        folder.mkdir()
        args, expect = spec["build"](folder)
        task = {"id": spec["id"], "category": spec["category"], "title": spec["title"],
                "request": spec["request"], "args": args, "expect": expect}
        write_json(folder / "task.json", task)
    print(f"Wrote {len(CASE_LIST)} cases to {CASES}")


if __name__ == "__main__":
    main()
