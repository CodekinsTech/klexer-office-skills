#!/usr/bin/env python3
"""Run the PPTX v2 benchmark through the CLI and write results.md.

For each cases/<id>/task.json: copy the case folder to _work/<id>/, run the
CLI there, then compare what happened with the expected result in task.json:
exit code, stderr text, JSON on stdout, output files, folder contents and the
input files (which must be unchanged). Output decks are opened with python-pptx
directly, not through the tool.

Usage (use the venv that has python-pptx):
    python run_benchmark.py [--only TEXT] [--tool PATH] [--python PATH]
"""
import argparse
import difflib
import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
CASES = HERE / "cases"
WORK = HERE / "_work"
RESULTS = HERE / "results.md"
DEFAULT_TOOL = Path(__file__).resolve().parents[1] / "scripts" / "pptx_tool.py"
DEFAULT_PYTHON = Path(sys.executable)
TIMEOUT_S = 120

try:
    from pptx import Presentation, __version__ as PPTX_VERSION
    from pptx.enum.dml import MSO_FILL
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    print("python-pptx is needed to read the output decks. Run this with the benchmark venv.",
          file=sys.stderr)
    sys.exit(1)


# ------------------------------------------------------------ comparison

def canon(value):
    """Make values comparable: integral floats become ints, lists and dicts recurse."""
    if isinstance(value, bool) or value is None or isinstance(value, str):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, (list, tuple)):
        return [canon(v) for v in value]
    if isinstance(value, dict):
        return {str(k): canon(v) for k, v in value.items()}
    return value


def as_text(value):
    if isinstance(value, str):
        return value
    return json.dumps(canon(value), indent=2, ensure_ascii=False, sort_keys=True)


def norm_ws(text):
    return " ".join((text or "").split())


class Checks:
    def __init__(self):
        self.items = []  # (ok, label, expected, actual)

    def eq(self, label, expected, actual):
        same = (json.dumps(canon(expected), sort_keys=True, ensure_ascii=False)
                == json.dumps(canon(actual), sort_keys=True, ensure_ascii=False))
        self.items.append((same, label, expected, actual))

    def has(self, label, needle, haystack):
        self.items.append((needle in haystack, label, f"contains {needle!r}", haystack))

    def failed(self):
        return [item for item in self.items if not item[0]]


# ------------------------------------------------------------ deck reading

def _walk(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(shape.shapes)


def _chart_info(chart):
    info = {"type": chart.chart_type.name}
    plot = chart.plots[0]
    try:
        info["categories"] = [str(c) for c in plot.categories]
    except Exception:
        info["categories"] = None
    info["series"] = []
    info["series_fill"] = []
    for series in plot.series:
        info["series"].append([series.name, list(series.values)])
        try:
            fill = series.format.fill
            info["series_fill"].append(
                str(fill.fore_color.rgb) if fill.type == MSO_FILL.SOLID else None)
        except Exception:
            info["series_fill"].append(None)
    return info


def read_deck(path):
    """Read a deck with python-pptx. Never creates a notes page."""
    prs = Presentation(str(path))
    slides = []
    for number, slide in enumerate(prs.slides, start=1):
        has_notes_page = slide.has_notes_slide
        notes = ""
        if has_notes_page:
            ntf = slide.notes_slide.notes_text_frame
            notes = ntf.text if ntf is not None else ""
        title_shape = slide.shapes.title
        title_id = title_shape.shape_id if title_shape is not None else None
        rec = {"number": number,
               "title": title_shape.text_frame.text if title_shape is not None else None,
               "body": [], "texts": [], "tables": [], "chart": None, "pictures": [],
               "notes": notes, "has_notes_page": has_notes_page}
        for shape in _walk(slide.shapes):
            is_title = title_id is not None and shape.shape_id == title_id
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if para.text.strip():
                        rec["texts"].append(para.text)
                        if not is_title:
                            rec["body"].append([para.text, para.level + 1])
            if shape.has_table:
                rows = [[cell.text for cell in row.cells] for row in shape.table.rows]
                rec["tables"].append(rows)
                rec["texts"].extend(c for row in rows for c in row if c.strip())
            if shape.has_chart and rec["chart"] is None:
                rec["chart"] = _chart_info(shape.chart)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                rec["pictures"].append({
                    "sha256": hashlib.sha256(shape.image.blob).hexdigest(),
                    "name": shape.name,
                    "left": int(shape.left), "top": int(shape.top),
                    "width": int(shape.width), "height": int(shape.height)})
        rec["table"] = rec["tables"][0] if rec["tables"] else None
        slides.append(rec)
    return {"slide_count": len(slides), "slides": slides}


def package_hit(path, needle):
    """Return the zip member whose name or bytes contain needle, or None."""
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if needle in name or needle.encode("utf-8") in z.read(name):
                return name
    return None


def check_deck(checks, prefix, path, want):
    try:
        deck = read_deck(path)
    except Exception as exc:  # the output is not a readable deck
        checks.eq(f"{prefix} opens with python-pptx", "opens",
                  f"{type(exc).__name__}: {exc}")
        return
    slides = deck["slides"]

    def slide(n):
        n = int(n)
        return slides[n - 1] if 1 <= n <= len(slides) else None

    if "slide_count" in want:
        checks.eq(f"{prefix} slide count", want["slide_count"], len(slides))
    if "titles" in want:
        checks.eq(f"{prefix} titles", want["titles"], [s["title"] for s in slides])
    if "notes_pages" in want:
        checks.eq(f"{prefix} slides with a notes page", want["notes_pages"],
                  sum(1 for s in slides if s["has_notes_page"]))
    for n, value in want.get("body", {}).items():
        s = slide(n)
        checks.eq(f"{prefix} slide {n} body", value, s["body"] if s else "<no slide>")
    for n, value in want.get("notes", {}).items():
        s = slide(n)
        checks.eq(f"{prefix} slide {n} notes", norm_ws(value),
                  norm_ws(s["notes"]) if s else "<no slide>")
    for n, value in want.get("table", {}).items():
        s = slide(n)
        checks.eq(f"{prefix} slide {n} table", value, s["table"] if s else "<no slide>")
    for n, value in want.get("chart", {}).items():
        s = slide(n)
        got = s["chart"] if s else None
        if got is None:
            checks.eq(f"{prefix} slide {n} chart", "chart present", "no chart")
        else:
            checks.eq(f"{prefix} slide {n} chart", value,
                      {k: got.get(k) for k in value})
    for n, value in want.get("pictures", {}).items():
        s = slide(n)
        got = s["pictures"] if s else []
        if len(got) != len(value):
            checks.eq(f"{prefix} slide {n} picture count", len(value), len(got))
        else:
            sub = [{k: g.get(k) for k in w} for w, g in zip(value, got)]
            checks.eq(f"{prefix} slide {n} pictures", value, sub)
    all_text = "\n".join(
        "\n".join(s["texts"] + ([s["notes"]] if s["notes"] else [])) for s in slides)
    for phrase, count in want.get("count_in_deck", {}).items():
        checks.eq(f"{prefix} occurrences of {phrase!r}", count, all_text.count(phrase))
    for phrase in want.get("all_text_absent", []):
        checks.eq(f"{prefix} no {phrase!r} left in text", "absent",
                  "absent" if phrase not in all_text else "present")
    for needle in want.get("package_absent", []):
        try:
            hit = package_hit(path, needle)
        except Exception as exc:
            hit = f"could not scan package: {exc}"
        checks.eq(f"{prefix} {needle!r} absent from every package part", "absent",
                  "absent" if hit is None else f"found in {hit}")


def json_get(obj, dotted):
    cur = obj
    for part in dotted.split("."):
        if isinstance(cur, list):
            cur = cur[int(part)]
        else:
            cur = cur[part]
    return cur


def sha256_file(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


# ------------------------------------------------------------ one case

def run_case(task, case_dir, tool, python):
    """Return (checks, stdout, stderr, returncode, note)."""
    work = WORK / task["id"]
    if work.exists():
        shutil.rmtree(work)
    shutil.copytree(case_dir, work, ignore=shutil.ignore_patterns("task.json"))
    inputs = {p.name: sha256_file(p) for p in work.iterdir() if p.is_file()}

    exp = task["expect"]
    checks = Checks()
    note = ""
    try:
        proc = subprocess.run([str(python), str(tool), *task["args"]], cwd=str(work),
                              capture_output=True, timeout=TIMEOUT_S)
        rc = proc.returncode
        out = proc.stdout.decode("utf-8", "replace")
        err = proc.stderr.decode("utf-8", "replace")
    except subprocess.TimeoutExpired as exc:
        rc = None
        out = (exc.stdout or b"").decode("utf-8", "replace")
        err = (exc.stderr or b"").decode("utf-8", "replace")
        note = f"timed out after {TIMEOUT_S} s"
        checks.eq("finished within the time limit", True, False)
        return checks, out, err, rc, note

    checks.eq("exit code", exp["exit"], rc)
    for needle in exp.get("stderr_has", []):
        checks.has("stderr", needle, err)
    if exp.get("stderr_nonempty"):
        checks.eq("stderr has a message", True, bool(err.strip()))

    if "json" in exp:
        try:
            obj = json.loads(out)
        except ValueError:
            checks.eq("stdout is one JSON object", "JSON object", out[:300])
            obj = None
        if obj is not None:
            for key, want in exp["json"].items():
                # Key forms: "a.b.0.c" exact; "len:a" length; "~a" case-insensitive
                # substring; "all:list.field" every item's field equals want.
                mode, path = None, key
                if key.startswith("len:"):
                    mode, path = "len", key[4:]
                elif key.startswith("~"):
                    mode, path = "ci", key[1:]
                elif key.startswith("all:"):
                    mode, path = "all", key[4:]
                try:
                    if mode == "all":
                        list_path, field = path.rsplit(".", 1)
                        got = [item.get(field) for item in json_get(obj, list_path)]
                    else:
                        got = json_get(obj, path)
                except (KeyError, IndexError, ValueError, TypeError, AttributeError):
                    got = "<missing>"
                if mode == "len":
                    got = len(got) if isinstance(got, (list, dict, str)) else "<no length>"
                    checks.eq(f"json {key}", want, got)
                elif mode == "ci":
                    ok = isinstance(got, str) and want.lower() in got.lower()
                    checks.items.append((ok, f"json {key}",
                                         f"contains {want!r} (case-insensitive)", got))
                elif mode == "all":
                    expected_list = [want] * len(got) if isinstance(got, list) else want
                    checks.eq(f"json {key} (every item)", expected_list, got)
                else:
                    checks.eq(f"json {key}", want, got)

    should_exist_default = exp["exit"] == 0
    for rel, want in exp.get("outputs", {}).items():
        path = work / rel
        exists = path.is_file()
        checks.eq(f"output {rel} written", should_exist_default, exists)
        if should_exist_default and exists and isinstance(want, dict):
            check_deck(checks, f"output {rel}", path, want)

    for rel, count in exp.get("dir_files", {}).items():
        folder = work / rel
        n = len([f for f in folder.iterdir() if f.is_file()]) if folder.is_dir() else 0
        checks.eq(f"files in folder {rel}", count, n)

    if "input_deck" in exp:
        spec = dict(exp["input_deck"])
        name = spec.pop("file")
        check_deck(checks, f"input {name}", work / name, spec)

    for name, digest in inputs.items():
        now = sha256_file(work / name) if (work / name).is_file() else None
        checks.eq(f"input {name} left unchanged", digest, now)

    return checks, out, err, rc, note


# ------------------------------------------------------------ report

def diff_block(expected, actual):
    lines = list(difflib.unified_diff(as_text(expected).splitlines(),
                                      as_text(actual).splitlines(),
                                      "expected", "actual", lineterm=""))
    return "\n".join(lines) if lines else "(no line diff)"


def excerpt(text, limit=1500):
    text = text.strip()
    if not text:
        return "(empty)"
    return text if len(text) <= limit else text[:limit] + "\n... (truncated)"


def write_results(records, tool, python, started):
    total_pass = sum(1 for r in records if r["runnable"] and r["passed"])
    total_run = sum(1 for r in records if r["runnable"])
    cats = OrderedDict()
    for r in records:
        c = cats.setdefault(r["category"], {"pass": 0, "run": 0, "fails": []})
        if r["runnable"]:
            c["run"] += 1
            if r["passed"]:
                c["pass"] += 1
            else:
                c["fails"].append(r["id"])

    out = []
    out.append("# PPTX v2 accuracy benchmark results\n")
    out.append(f"- Run: {started:%Y-%m-%d %H:%M}")
    out.append(f"- Tool: `{tool}`")
    out.append(f"- Python: `{python}` (python-pptx {PPTX_VERSION})")
    out.append(f"- Tasks: {len(records)}, runnable: {total_run}")
    pct = (100.0 * total_pass / total_run) if total_run else 0.0
    out.append(f"- **Score: {total_pass}/{total_run} passed ({pct:.0f}%)**\n")

    out.append("## Score by category\n")
    out.append("| Category | Passed / runnable | Failed tasks |")
    out.append("|---|---|---|")
    for name, c in sorted(cats.items(), key=lambda kv: (len(kv[1]["fails"]) * -1, kv[0])):
        failed = ", ".join(c["fails"]) if c["fails"] else "none"
        out.append(f"| {name} | {c['pass']}/{c['run']} | {failed} |")
    out.append("")

    ranked = [(name, len(c["fails"])) for name, c in cats.items() if c["fails"]]
    ranked.sort(key=lambda kv: -kv[1])
    out.append("## Top failure categories\n")
    if ranked:
        for name, n in ranked:
            out.append(f"- {name}: {n} failed")
    else:
        out.append("- none")
    out.append("")

    out.append("## Per-task results\n")
    out.append("| # | Task | Category | Result | Exit (expected / actual) |")
    out.append("|---|---|---|---|---|")
    for i, r in enumerate(records, 1):
        result = "NOT RUN" if not r["runnable"] else ("PASS" if r["passed"] else "FAIL")
        actual = "-" if r["rc"] is None else str(r["rc"])
        out.append(f"| {i} | `{r['id']}` {r['title']} | {r['category']} | {result} | "
                   f"{r['exit_expected']} / {actual} |")
    out.append("")

    failures = [r for r in records if r["runnable"] and not r["passed"]]
    out.append("## Failures\n")
    if not failures:
        out.append("No failures.\n")
    for r in failures:
        out.append(f"### `{r['id']}`: {r['title']}\n")
        out.append(f"- Category: {r['category']}")
        out.append(f"- Request: {r['request']}")
        out.append(f"- Command: `{' '.join(r['args'])}`")
        if r["note"]:
            out.append(f"- Note: {r['note']}")
        if r.get("exp_note"):
            out.append(f"- Expectation note: {r['exp_note']}")
        out.append("")
        for ok, label, expected, actual in r["failed_checks"]:
            out.append(f"**{label}**\n")
            out.append("```diff")
            out.append(diff_block(expected, actual))
            out.append("```\n")
        out.append("<details><summary>stdout and stderr</summary>\n")
        out.append("```text")
        out.append("stdout:\n" + excerpt(r["stdout"]))
        out.append("\nstderr:\n" + excerpt(r["stderr"]))
        out.append("```\n</details>\n")

    not_run = [r for r in records if not r["runnable"]]
    if not_run:
        out.append("## Not runnable\n")
        for r in not_run:
            out.append(f"- `{r['id']}`: {r['note']}")
        out.append("")

    RESULTS.write_text("\n".join(out) + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--only", default="", help="run only task ids containing this text")
    parser.add_argument("--tool", default=str(DEFAULT_TOOL), help="path to pptx_tool.py")
    parser.add_argument("--python", default=str(DEFAULT_PYTHON),
                        help="python that runs the CLI (the venv with its requirements)")
    opts = parser.parse_args()

    if not CASES.is_dir():
        print(f"No cases folder at {CASES}. Run make_tasks.py first.", file=sys.stderr)
        return 1
    tool = Path(opts.tool)
    python = Path(opts.python)
    started = datetime.now()

    records = []
    for case_dir in sorted(p for p in CASES.iterdir() if p.is_dir()):
        if opts.only and opts.only not in case_dir.name:
            continue
        task_file = case_dir / "task.json"
        try:
            task = json.loads(task_file.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            records.append({"id": case_dir.name, "title": "", "category": "?", "request": "",
                            "args": [], "runnable": False, "passed": False, "rc": None,
                            "exit_expected": "?", "failed_checks": [], "stdout": "",
                            "stderr": "", "note": f"task.json unreadable: {exc}"})
            continue
        if not tool.is_file() or not python.is_file():
            records.append({"id": task["id"], "title": task["title"],
                            "category": task["category"], "request": task["request"],
                            "args": task["args"], "runnable": False, "passed": False,
                            "rc": None, "exit_expected": task["expect"]["exit"],
                            "failed_checks": [], "stdout": "", "stderr": "",
                            "note": "CLI or python path not found"})
            continue
        try:
            checks, out, err, rc, note = run_case(task, case_dir, tool, python)
        except OSError as exc:
            records.append({"id": task["id"], "title": task["title"],
                            "category": task["category"], "request": task["request"],
                            "args": task["args"], "runnable": False, "passed": False,
                            "rc": None, "exit_expected": task["expect"]["exit"],
                            "failed_checks": [], "stdout": "", "stderr": "",
                            "note": f"could not start the CLI: {exc}"})
            continue
        failed = checks.failed()
        records.append({"id": task["id"], "title": task["title"],
                        "category": task["category"], "request": task["request"],
                        "args": task["args"], "runnable": True, "passed": not failed,
                        "rc": rc, "exit_expected": task["expect"]["exit"],
                        "failed_checks": failed, "stdout": out, "stderr": err,
                        "note": note, "exp_note": task["expect"].get("note", "")})
        mark = "PASS" if not failed else "FAIL"
        print(f"{mark}  {task['id']}  ({len(checks.items) - len(failed)}/{len(checks.items)} checks)")

    write_results(records, tool, python, started)
    total_pass = sum(1 for r in records if r["runnable"] and r["passed"])
    total_run = sum(1 for r in records if r["runnable"])
    print(f"Score: {total_pass}/{total_run}. Wrote {RESULTS}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
