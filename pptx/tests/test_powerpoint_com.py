"""PowerPoint checks for the pptx tool's written decks, through Microsoft PowerPoint COM automation.

For each command that writes a deck (create, replace, add, delete, reorder,
duplicate, split, replace-image, chart-data), the output is opened in PowerPoint
with WithWindow false. The test passes only if PowerPoint opens the file without
a repair prompt, the slide count is the one the tool reports, and every slide's
title, text, table text, chart categories, series values and speaker notes match
what the tool's own `read --json` reports for the same file.

PowerPoint runs in a child process (this file, with --readout). A repair prompt
or any other modal dialog blocks COM calls, so the parent waits with a timeout
and fails the test instead of hanging. On a timeout the parent stops only the
POWERPNT.EXE instances that started during the test; instances that were already
running are left alone. Each child opens its own instance (DispatchEx), closes
the presentation and quits that instance in a finally block.

The tests skip automatically when Windows, the PowerPoint COM server, or pywin32
is not available.

Run from the skill folder:  python -m pytest tests/test_powerpoint_com.py -q
"""
import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from PIL import Image as PILImage
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

TOOL = Path(__file__).resolve().parents[1] / "scripts" / "pptx_tool.py"
THIS_FILE = Path(__file__).resolve()
READOUT_TIMEOUT_S = 90
PROBE_TIMEOUT_S = 120
PP_ALERTS_ALL = 2          # ppAlertsAll: a repair prompt is shown, so it blocks instead of being hidden
MSO_GROUP = 6
MSO_PICTURE = 13
MSO_PLACEHOLDER = 14
PP_PLACEHOLDER_TITLE = (1, 3)   # ppPlaceholderTitle, ppPlaceholderCenterTitle
PP_PLACEHOLDER_BODY = 2
XL_COLUMN_CLUSTERED = 51


# ------------------------------------------------------------------ PowerPoint side (child process)

def _norm(text):
    return text.replace("\r", "\n").strip()


def _walk(shapes):
    """Shapes in stacking order, groups expanded (as the tool does)."""
    for index in range(1, shapes.Count + 1):
        shape = shapes.Item(index)
        if shape.Type == MSO_GROUP:
            yield from _walk(shape.GroupItems)
        else:
            yield shape


def _chart_readout(chart):
    series, categories = [], []
    for index in range(1, chart.SeriesCollection().Count + 1):
        item = chart.SeriesCollection(index)
        if not categories:
            categories = [str(x) for x in item.XValues]
        series.append({"name": item.Name, "values": list(item.Values)})
    return {"chart_type": chart.ChartType, "categories": categories, "series": series}


def _slide_readout(slide):
    title, texts, tables, charts, pictures = "", [], [], [], []
    for shape in _walk(slide.Shapes):
        if shape.Type == MSO_PLACEHOLDER and shape.PlaceholderFormat.Type in PP_PLACEHOLDER_TITLE and not title:
            title = _norm(shape.TextFrame.TextRange.Text)
        if shape.Type == MSO_PICTURE:
            pictures.append([shape.Left, shape.Top, shape.Width, shape.Height])
        if shape.HasTable:
            table = shape.Table
            tables.append([
                [_norm(table.Cell(r, c).Shape.TextFrame.TextRange.Text) for c in range(1, table.Columns.Count + 1)]
                for r in range(1, table.Rows.Count + 1)
            ])
        if shape.HasChart:
            charts.append(_chart_readout(shape.Chart))
        if shape.HasTextFrame and shape.TextFrame.HasText:
            text = _norm(shape.TextFrame.TextRange.Text)
            if text:
                texts.append(text)
    notes = ""
    notes_shapes = slide.NotesPage.Shapes
    for index in range(1, notes_shapes.Count + 1):
        shape = notes_shapes.Item(index)
        if shape.Type == MSO_PLACEHOLDER and shape.PlaceholderFormat.Type == PP_PLACEHOLDER_BODY:
            notes = _norm(shape.TextFrame.TextRange.Text)
    return {"title": title, "texts": texts, "tables": tables, "charts": charts,
            "notes": notes, "pictures": pictures}


def _readout(path):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.DisplayAlerts = PP_ALERTS_ALL
        pres = app.Presentations.Open(str(path), True, False, False)
        try:
            slides = [_slide_readout(pres.Slides.Item(i)) for i in range(1, pres.Slides.Count + 1)]
            return {"slide_count": pres.Slides.Count, "slides": slides}
        finally:
            pres.Close()
    finally:
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()


def _probe():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        return app.Version
    finally:
        app.Quit()
        pythoncom.CoUninitialize()


def _make_deck(out, picture):
    """A deck authored by PowerPoint itself: notes, a two-paragraph body, a table, a picture and sections."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.DisplayAlerts = 1
    pres = None
    try:
        pres = app.Presentations.Add(0)
        layout = pres.SlideMaster.CustomLayouts.Item(2)   # "Title and Content" in the default theme
        for number, title in enumerate(["Alpha", "Beta", "Gamma", "Delta"], start=1):
            slide = pres.Slides.AddSlide(number, layout)
            slide.Shapes.Item(1).TextFrame.TextRange.Text = title
            slide.Shapes.Item(2).TextFrame.TextRange.Text = f"{title} body\rsecond paragraph"
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = f"{title} notes"
        pres.Slides.Item(1).Shapes.AddPicture(picture, 0, -1, 100, 300, 200, 150)
        table = pres.Slides.Item(3).Shapes.AddTable(2, 2, 50, 300, 400, 100).Table
        table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Region"
        table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Total"
        table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "North"
        table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "42"
        pres.SectionProperties.AddBeforeSlide(1, "First section")
        pres.SectionProperties.AddBeforeSlide(3, "Second section")
        pres.SaveAs(str(out), 24)   # ppSaveAsOpenXMLPresentation
    finally:
        if pres is not None:
            pres.Close()
        app.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--readout":
        print(json.dumps(_readout(sys.argv[2])))
        sys.exit(0)
    if len(sys.argv) == 4 and sys.argv[1] == "--make-deck":
        _make_deck(sys.argv[2], sys.argv[3])
        sys.exit(0)
    if len(sys.argv) == 2 and sys.argv[1] == "--probe":
        print(_probe())
        sys.exit(0)
    sys.exit("usage: test_powerpoint_com.py --readout FILE | --make-deck OUT PNG | --probe")


# ------------------------------------------------------------------ parent side (pytest)

def _powerpoint_pids():
    listing = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True, text=True,
    ).stdout
    pids = set()
    for line in listing.splitlines():
        parts = line.strip().strip('"').split('","')
        if len(parts) >= 2 and parts[0].upper() == "POWERPNT.EXE" and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


@pytest.fixture(scope="module", autouse=True)
def powerpoint_available():
    if sys.platform != "win32":
        pytest.skip("PowerPoint COM tests run on Windows only")
    try:
        proc = subprocess.run([sys.executable, str(THIS_FILE), "--probe"], capture_output=True,
                              text=True, timeout=PROBE_TIMEOUT_S)
    except subprocess.TimeoutExpired:
        pytest.skip("PowerPoint did not start in time")
    if proc.returncode != 0:
        pytest.skip("PowerPoint COM is not available: " + (proc.stderr.strip().splitlines() or ["no detail"])[-1])


def make_powerpoint_deck(out: Path, picture: Path) -> Path:
    """Have PowerPoint author a deck (child process, same timeout and cleanup rules as readouts)."""
    baseline = _powerpoint_pids()
    try:
        proc = subprocess.run([sys.executable, str(THIS_FILE), "--make-deck", str(out), str(picture)],
                              capture_output=True, text=True, encoding="utf-8", timeout=READOUT_TIMEOUT_S)
    except subprocess.TimeoutExpired:
        for pid in _powerpoint_pids() - baseline:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
        pytest.fail("PowerPoint did not finish authoring the test deck in time.")
    if proc.returncode != 0:
        detail = (proc.stderr.strip().splitlines() or ["no detail"])[-1]
        pytest.fail(f"PowerPoint could not author the test deck: {detail}")
    return out


def powerpoint_readout(path: Path) -> dict:
    """Open the file in PowerPoint (in a child process) and return what it shows."""
    baseline = _powerpoint_pids()
    try:
        proc = subprocess.run([sys.executable, str(THIS_FILE), "--readout", str(path)],
                              capture_output=True, text=True, encoding="utf-8",
                              timeout=READOUT_TIMEOUT_S)
    except subprocess.TimeoutExpired:
        for pid in _powerpoint_pids() - baseline:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
        pytest.fail(f"PowerPoint did not finish opening {path.name} within {READOUT_TIMEOUT_S} s. "
                    "A dialog such as a repair prompt is likely waiting for input.")
    if proc.returncode != 0:
        detail = (proc.stderr.strip().splitlines() or ["no detail"])[-1]
        pytest.fail(f"PowerPoint could not open {path.name}: {detail}")
    return json.loads(proc.stdout.strip().splitlines()[-1])


def tool(*args):
    return subprocess.run([sys.executable, str(TOOL), *[str(a) for a in args]],
                          capture_output=True, text=True, encoding="utf-8")


def tool_json(*args):
    result = tool(*args, "--json")
    assert result.returncode == 0, result.stderr
    return json.loads(result.stdout)


def tool_read(path: Path) -> dict:
    return tool_json("read", path)


def _tool_slides(read_json):
    return [
        {
            "title": item["title"],
            "texts": item["texts"],
            "tables": item["tables"],
            "charts": [{"categories": c["categories"], "series": c["series"]} for c in item["charts"]],
            "notes": item["notes"],
        }
        for item in read_json["slides"]
    ]


def _com_slides(readout):
    return [
        {
            "title": slide["title"],
            "texts": slide["texts"],
            "tables": slide["tables"],
            "charts": [{"categories": c["categories"], "series": c["series"]} for c in slide["charts"]],
            "notes": slide["notes"],
        }
        for slide in readout["slides"]
    ]


def assert_powerpoint_matches_tool(path: Path, expected_count: int = None) -> dict:
    """PowerPoint opens the file cleanly and shows what the tool reports for it."""
    readout = powerpoint_readout(path)
    reported = tool_read(path)
    assert readout["slide_count"] == reported["slide_count"], (
        f"PowerPoint shows {readout['slide_count']} slides, the tool reports {reported['slide_count']}")
    if expected_count is not None:
        assert readout["slide_count"] == expected_count
    assert _com_slides(readout) == _tool_slides(reported), f"text differs in PowerPoint for {path.name}"
    return readout


def make_png(path, colour, size=100):
    PILImage.new("RGB", (size, size), colour).save(path)
    return path


def add_text_slide(prs, title, bullets=(), notes=None, layout=1):
    slide = prs.slides.add_slide(prs.slide_layouts[layout])
    slide.shapes.title.text_frame.text = title
    if bullets:
        body = slide.placeholders[1].text_frame
        for index, (text, level) in enumerate(bullets):
            paragraph = body.paragraphs[0] if index == 0 else body.add_paragraph()
            paragraph.text = text
            paragraph.level = level - 1
    if notes is not None:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


def titled_deck(path, titles):
    prs = Presentation()
    for title in titles:
        add_text_slide(prs, title, [(f"{title} point", 1)], notes=f"{title} note")
    prs.save(str(path))
    return path


# ------------------------------------------------------------------ create

CREATE_OUTLINE = (
    "# Intro\n"
    "- one\n"
    "  - sub\n"
    "Notes: hello\n"
    "\n"
    "# Café plans\n"
    "- naïve façade ✓\n"
    "- ภาษาไทย\n"
    "Notes: second line\n"
    "Notes: third line\n"
    "\n"
    "# End\n"
    "- bye\n"
)


def test_com_create_opens_and_matches(tmp_path):
    outline = tmp_path / "outline.txt"
    outline.write_text(CREATE_OUTLINE, encoding="utf-8")
    out = tmp_path / "made.pptx"
    created = tool_json("create", outline, "--out", out)
    assert created["slides"] == 3
    readout = assert_powerpoint_matches_tool(out, expected_count=3)
    titles = [s["title"] for s in readout["slides"]]
    assert titles == ["Intro", "Café plans", "End"]
    assert readout["slides"][0]["notes"] == "hello"
    assert readout["slides"][1]["notes"] == "second line\nthird line"
    assert "ภาษาไทย" in readout["slides"][1]["texts"][1]


# ------------------------------------------------------------------ replace

def test_com_replace_opens_and_matches(tmp_path):
    prs = Presentation()
    slide1 = add_text_slide(prs, "ACME launch", [("Split ", 1)])
    # Split-run occurrence: "AC" and "ME" in separate runs of one paragraph.
    paragraph = slide1.placeholders[1].text_frame.paragraphs[0]
    paragraph.runs[0].text = "AC"
    run = paragraph.add_run()
    run.text = "ME deal"
    slide2 = add_text_slide(prs, "Plan", [("Partner", 1)], notes="ACME notes")
    table = slide2.shapes.add_table(2, 2, Inches(1), Inches(3), Inches(4), Inches(1)).table
    table.cell(0, 0).text = "ACME Ltd"
    table.cell(0, 1).text = "Q1"
    table.cell(1, 0).text = "x"
    table.cell(1, 1).text = "y"
    src = tmp_path / "deck.pptx"
    prs.save(str(src))

    result = tool_json("replace", src, "--find", "ACME", "--replace", "Globex")
    assert result["replacements"] == 4
    assert result["notes_replacements"] == 1
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=2)
    everything = json.dumps(readout, ensure_ascii=False)
    assert "ACME" not in everything
    assert readout["slides"][0]["title"] == "Globex launch"
    assert readout["slides"][1]["tables"][0][0][0] == "Globex Ltd"
    assert readout["slides"][1]["notes"] == "Globex notes"


# ------------------------------------------------------------------ add

def test_com_add_opens_and_matches(tmp_path):
    src = titled_deck(tmp_path / "deck.pptx", ["A", "B"])
    result = tool_json("add", src, "--at", "2", "--title", "New slide",
                       "--body", "- point", "--body", "  - sub point", "--notes", "added note")
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=3)
    new = readout["slides"][1]
    assert new["title"] == "New slide"
    assert new["texts"] == ["New slide", "point\nsub point"]
    assert new["notes"] == "added note"


# ------------------------------------------------------------------ delete

def test_com_delete_opens_and_matches(tmp_path):
    src = titled_deck(tmp_path / "deck.pptx", ["One", "Two", "Three"])
    result = tool_json("delete", src, "--slides", "2")
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=2)
    assert [s["title"] for s in readout["slides"]] == ["One", "Three"]


# ------------------------------------------------------------------ reorder

def test_com_reorder_opens_and_matches(tmp_path):
    src = titled_deck(tmp_path / "deck.pptx", ["A", "B", "C"])
    result = tool_json("reorder", src, "--order", "3,1,2")
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=3)
    assert [s["title"] for s in readout["slides"]] == ["C", "A", "B"]


# ------------------------------------------------------------------ duplicate

def test_com_duplicate_with_picture_opens_and_matches(tmp_path):
    prs = Presentation()
    slide = add_text_slide(prs, "Pictured", [("Body text", 1)], notes="copy me")
    png = make_png(tmp_path / "red.png", (255, 0, 0))
    slide.shapes.add_picture(str(png), Inches(1), Inches(3), Inches(2), Inches(2))
    add_text_slide(prs, "Other", [("Second", 1)])
    src = tmp_path / "deck.pptx"
    prs.save(str(src))

    result = tool_json("duplicate", src, "--slide", "1")
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=3)
    original, copy = readout["slides"][0], readout["slides"][1]
    assert copy["title"] == "Pictured"
    assert copy["notes"] == "copy me"
    assert len(copy["pictures"]) == 1
    assert copy["pictures"] == original["pictures"]


# ------------------------------------------------------------------ split

def test_com_split_every_part_opens_and_matches(tmp_path):
    titles = ["S1", "S2", "S3", "S4", "S5"]
    src = titled_deck(tmp_path / "deck.pptx", titles)
    result = tool_json("split", src, "--every", "2", "--out-dir", tmp_path / "parts")
    assert [p["slides"] for p in result["parts"]] == [2, 2, 1]
    expected_titles = [titles[0:2], titles[2:4], titles[4:5]]
    for part, wanted in zip(result["parts"], expected_titles):
        readout = assert_powerpoint_matches_tool(Path(part["file"]), expected_count=part["slides"])
        assert [s["title"] for s in readout["slides"]] == wanted


# ------------------------------------------------------------------ replace-image

def test_com_replace_image_opens_and_matches(tmp_path):
    prs = Presentation()
    slide = add_text_slide(prs, "Logo slide")
    red = make_png(tmp_path / "red.png", (255, 0, 0))
    slide.shapes.add_picture(str(red), Inches(1), Inches(1), Inches(3), Inches(2))
    src = tmp_path / "deck.pptx"
    prs.save(str(src))
    blue = make_png(tmp_path / "blue.png", (0, 0, 255), size=50)

    result = tool_json("replace-image", src, "--slide", "1", "--image", blue)
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=1)
    pictures = readout["slides"][0]["pictures"]
    assert len(pictures) == 1
    # Positions are in points: 1 inch = 72 points.
    assert pictures[0] == [72.0, 72.0, 216.0, 144.0]


# ------------------------------------------------------------------ chart-data

def test_com_chart_data_opens_and_matches(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text_frame.text = "Sales"
    data = CategoryChartData()
    data.categories = ["Q1", "Q2"]
    data.add_series("Sales", (10, 12))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(6), Inches(4), data)
    src = tmp_path / "deck.pptx"
    prs.save(str(src))
    new_data = tmp_path / "new.json"
    new_data.write_text(json.dumps({"categories": ["Q1", "Q2", "Q3"],
                                    "series": [{"name": "Sales", "values": [5, 6, 7]}]}), encoding="utf-8")

    result = tool_json("chart-data", src, "--slide", "1", "--data", new_data)
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=1)
    chart = readout["slides"][0]["charts"][0]
    assert chart["categories"] == ["Q1", "Q2", "Q3"]
    assert chart["series"] == [{"name": "Sales", "values": [5, 6, 7]}]
    assert chart["chart_type"] == XL_COLUMN_CLUSTERED


# ------------------------------------------------------------------ PowerPoint-authored source decks

@pytest.fixture(scope="module")
def ppt_source(tmp_path_factory):
    """One deck written by PowerPoint, shared by the tests below (each test copies it first)."""
    folder = tmp_path_factory.mktemp("ppt_source")
    picture = make_png(folder / "red.png", (255, 0, 0))
    deck = make_powerpoint_deck(folder / "powerpoint_made.pptx", picture)
    readout = powerpoint_readout(deck)
    assert readout["slide_count"] == 4
    assert [s["title"] for s in readout["slides"]] == ["Alpha", "Beta", "Gamma", "Delta"]
    return deck


def copy_of(source: Path, tmp_path: Path) -> Path:
    target = tmp_path / source.name
    target.write_bytes(source.read_bytes())
    return target


def test_powerpoint_made_deck_reads_the_same_in_tool(ppt_source):
    assert_powerpoint_matches_tool(ppt_source, expected_count=4)


def test_com_duplicate_on_powerpoint_made_deck(ppt_source, tmp_path):
    src = copy_of(ppt_source, tmp_path)
    result = tool_json("duplicate", src, "--slide", "1")
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=5)
    copy = readout["slides"][1]
    assert copy["title"] == "Alpha"
    assert copy["texts"] == ["Alpha", "Alpha body\nsecond paragraph"]
    assert copy["notes"] == "Alpha notes"
    assert len(copy["pictures"]) == 1 and copy["pictures"] == readout["slides"][0]["pictures"]


def test_com_delete_on_powerpoint_made_deck(ppt_source, tmp_path):
    src = copy_of(ppt_source, tmp_path)
    result = tool_json("delete", src, "--slides", "2,4")
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=2)
    assert [s["title"] for s in readout["slides"]] == ["Alpha", "Gamma"]


def test_com_replace_image_on_powerpoint_made_deck(ppt_source, tmp_path):
    src = copy_of(ppt_source, tmp_path)
    blue = make_png(tmp_path / "blue.png", (0, 0, 255), size=50)
    result = tool_json("replace-image", src, "--slide", "1", "--image", blue)
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=4)
    assert len(readout["slides"][0]["pictures"]) == 1
    assert readout["slides"][0]["pictures"][0] == [100.0, 300.0, 200.0, 150.0]


def test_com_reorder_on_powerpoint_made_deck(ppt_source, tmp_path):
    src = copy_of(ppt_source, tmp_path)
    result = tool_json("reorder", src, "--order", "4,3,2,1")
    readout = assert_powerpoint_matches_tool(Path(result["output"]), expected_count=4)
    assert [s["title"] for s in readout["slides"]] == ["Delta", "Gamma", "Beta", "Alpha"]


# ------------------------------------------------------------------ detector control

def test_damaged_deck_is_flagged_not_passed(tmp_path):
    """Control: a package with a dangling picture relationship must not pass the PowerPoint check."""
    prs = Presentation()
    add_text_slide(prs, "Control")
    good = tmp_path / "good.pptx"
    prs.save(str(good))

    damaged = tmp_path / "damaged.pptx"
    with zipfile.ZipFile(good) as src, zipfile.ZipFile(damaged, "w", zipfile.ZIP_DEFLATED) as dst:
        for name in src.namelist():
            data = src.read(name)
            if name == "ppt/slides/slide1.xml":
                pic = (
                    "<p:pic><p:nvPicPr><p:cNvPr id='99' name='x'/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>"
                    "<p:blipFill><a:blip xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' "
                    "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' r:embed='rId99'/>"
                    "<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x='0' y='0'/>"
                    "<a:ext cx='914400' cy='914400'/></a:xfrm><a:prstGeom prst='rect'/></p:spPr></p:pic>"
                )
                data = data.decode("utf-8").replace("</p:spTree>", pic + "</p:spTree>").encode("utf-8")
            dst.writestr(name, data)
    with pytest.raises(pytest.fail.Exception):
        powerpoint_readout(damaged)
