#!/usr/bin/env python3
"""Run every benchmark task through the xlsx CLI and write results.md.

Run after make_tasks.py, with the benchmark venv:
    python run_benchmark.py
    python run_benchmark.py --only errors_div0_after_recalc

Each task runs in a fresh copy of its case folder. Every step's exit code and
checks are compared with the expected results in tasks.json. Output files are
read back with openpyxl, not with the tool.
"""

import argparse
import datetime as dt
import difflib
import json
import os
import shutil
import subprocess
import sys
from collections import OrderedDict
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
CASES = HERE / "cases"
WORK = HERE / "work"
TASKS_JSON = HERE / "tasks.json"
RESULTS_MD = HERE / "results.md"
DEFAULT_PYTHON = Path(sys.executable)
                      r"\59914a7a-0dc9-4ea2-93ba-658242a2ee94\scratchpad\venv-v2-xlsx\Scripts\python.exe")
DEFAULT_CLI = HERE.parent / "scripts" / "xlsx_tool.py"
TIMEOUT_SECONDS = 300
CATEGORY_ORDER = ["sheets", "read", "write", "format", "recalc", "refusal", "errors", "validate", "check-values"]


# ---------------------------------------------------------------- comparison

def normalize(value):
    """Dates from openpyxl become ISO strings, the same form the tool uses in JSON."""
    if isinstance(value, dt.datetime):
        return value.date().isoformat() if value.time() == dt.time(0) else value.isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    return value


def same(actual, expected):
    """Type-aware comparison. Booleans must be booleans. Dicts must have exactly the expected keys."""
    if isinstance(expected, bool) or isinstance(actual, bool):
        return type(actual) is type(expected) and actual == expected
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(actual - expected) < 1e-9
    if isinstance(expected, list):
        return (isinstance(actual, list) and len(actual) == len(expected)
                and all(same(a, e) for a, e in zip(actual, expected)))
    if isinstance(expected, dict):
        return (isinstance(actual, dict) and set(actual) == set(expected)
                and all(same(actual[k], expected[k]) for k in expected))
    return actual == expected


def sort_key(item):
    return json.dumps(item, sort_keys=True, default=str)


def pretty(obj):
    return json.dumps(obj, indent=2, ensure_ascii=False, sort_keys=True, default=str).splitlines()


def diff_of(expected, actual):
    lines = difflib.unified_diff(pretty(expected), pretty(actual), "expected", "actual", lineterm="")
    return "\n".join(lines)


def failure(what, expected, actual):
    return {"what": what, "expected": expected, "actual": actual,
            "diff": diff_of(expected, actual)}


def dig(data, path):
    current = data
    for key in path:
        current = current[key]
    return current


def read_cell(file_path, sheet, ref, mode):
    if mode == "format":
        return openpyxl.load_workbook(file_path)[sheet][ref].number_format
    wb = openpyxl.load_workbook(file_path, data_only=(mode == "value"))
    value = wb[sheet][ref].value
    if hasattr(value, "text"):  # array formula object
        value = value.text
    return normalize(value)


# ---------------------------------------------------------------- checks

def run_check(chk, cwd, stdout, stderr):
    kind = chk["check"]
    label = kind

    if kind in ("json", "json_contains", "json_truthy"):
        path = chk["path"]
        label = f"stdout JSON {'.'.join(map(str, path))}"
        try:
            data = json.loads(stdout)
        except ValueError:
            return [failure(f"{label}: stdout is not JSON", "a JSON object", stdout[:800])]
        try:
            actual = dig(data, path)
        except (KeyError, IndexError, TypeError):
            return [failure(f"{label}: key missing", chk.get("equals", chk.get("text", "present")), "<missing>")]
        if kind == "json":
            expected = chk["equals"]
            if chk.get("unordered") and isinstance(actual, list) and isinstance(expected, list):
                actual_cmp = sorted(actual, key=sort_key)
                expected_cmp = sorted(expected, key=sort_key)
            else:
                actual_cmp, expected_cmp = actual, expected
            if not same(actual_cmp, expected_cmp):
                return [failure(label, expected_cmp, actual_cmp)]
            return []
        if kind == "json_contains":
            if chk["text"] not in json.dumps(actual, ensure_ascii=False):
                return [failure(f"{label} contains {chk['text']!r}", chk["text"], actual)]
            return []
        if not actual:
            return [failure(f"{label} is truthy", True, actual)]
        return []

    if kind in ("stdout_contains", "stderr_contains"):
        stream = stdout if kind == "stdout_contains" else stderr
        if chk["text"] not in stream:
            return [failure(f"{kind}", chk["text"], stream.strip()[:800] or "<empty>")]
        return []

    if kind == "file_exists":
        exists = (cwd / chk["file"]).exists()
        if exists != chk["exists"]:
            return [failure(f"file {chk['file']} exists", chk["exists"], exists)]
        return []

    if kind == "cell":
        label = f"{chk['file']} {chk['sheet']}!{chk['cell']} ({chk['mode']})"
        try:
            actual = read_cell(cwd / chk["file"], chk["sheet"], chk["cell"], chk["mode"])
        except Exception as exc:  # missing file, missing sheet, bad file
            return [failure(label, chk["equals"], f"<could not read: {type(exc).__name__}: {exc}>")]
        if not same(actual, chk["equals"]):
            return [failure(label, chk["equals"], actual)]
        return []

    return [failure(f"unknown check kind {kind}", "known check", kind)]


# ---------------------------------------------------------------- running

def run_step(python, cli, args, cwd):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    try:
        return subprocess.run([str(python), str(cli), *args], cwd=cwd, capture_output=True,
                              text=True, encoding="utf-8", errors="replace",
                              timeout=TIMEOUT_SECONDS, env=env)
    except subprocess.TimeoutExpired:
        return None


def run_task(task, python, cli):
    work = WORK / task["id"]
    if work.exists():
        shutil.rmtree(work)
    shutil.copytree(CASES / task["id"], work)

    steps_report = []
    for number, st in enumerate(task["steps"], start=1):
        command = "xlsx_tool.py " + " ".join(st["args"])
        proc = run_step(python, cli, st["args"], work)
        failures = []
        if proc is None:
            failures.append(failure(f"step {number} finished", "within timeout",
                                    f"timed out after {TIMEOUT_SECONDS} seconds"))
            exit_actual = None
            stdout = stderr = ""
        else:
            exit_actual = proc.returncode
            stdout, stderr = proc.stdout, proc.stderr
            if exit_actual != st["exit"]:
                failures.append(failure(f"step {number} exit code", st["exit"], exit_actual))
            for chk in st["checks"]:
                failures.extend(run_check(chk, work, stdout, stderr))
        steps_report.append({"number": number, "command": command, "exit_expected": st["exit"],
                             "exit_actual": exit_actual, "stderr": stderr.strip()[:800],
                             "failures": failures})
    passed = all(not s["failures"] for s in steps_report)
    return {"id": task["id"], "category": task["category"], "request": task["request"],
            "passed": passed, "steps": steps_report}


# ---------------------------------------------------------------- report

def render(results, skipped, python, cli, started):
    runnable = [r for r in results]
    passed_total = sum(1 for r in runnable if r["passed"])
    score_pct = (100.0 * passed_total / len(runnable)) if runnable else 0.0

    by_category = OrderedDict((c, [0, 0]) for c in CATEGORY_ORDER)
    for r in runnable:
        by_category.setdefault(r["category"], [0, 0])
        by_category[r["category"]][1] += 1
        if r["passed"]:
            by_category[r["category"]][0] += 1

    failed = [r for r in runnable if not r["passed"]]
    fail_by_category = {}
    for r in failed:
        fail_by_category[r["category"]] = fail_by_category.get(r["category"], 0) + 1

    lines = []
    lines.append("# XLSX v2 skill accuracy benchmark: results")
    lines.append("")
    lines.append(f"Run started: {started.strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"CLI: `{cli}`")
    lines.append(f"Python: `{python}`")
    lines.append("")
    lines.append(f"**Score: {passed_total}/{len(runnable)} runnable tasks passed ({score_pct:.1f}%).**")
    if skipped:
        lines.append(f"{len(skipped)} task(s) could not be set up and are not in the score (see the end).")
    lines.append("")
    lines.append("## Score by category")
    lines.append("")
    lines.append("| Category | Passed | Runnable | Score |")
    lines.append("|---|---:|---:|---:|")
    for cat, (ok, total) in by_category.items():
        if total == 0:
            continue
        lines.append(f"| {cat} | {ok} | {total} | {100.0 * ok / total:.0f}% |")
    lines.append("")
    if fail_by_category:
        ranked = sorted(fail_by_category.items(), key=lambda kv: (-kv[1], kv[0]))
        lines.append("Failures by category: " + ", ".join(f"{c} ({n})" for c, n in ranked) + ".")
        lines.append("")
    else:
        lines.append("No failures.")
        lines.append("")

    lines.append("## Per-task results")
    lines.append("")
    lines.append("| # | Task | Category | Result | Failed step(s) |")
    lines.append("|---:|---|---|---|---|")
    for i, r in enumerate(runnable, start=1):
        bad = [str(s["number"]) for s in r["steps"] if s["failures"]]
        result = "PASS" if r["passed"] else "FAIL"
        lines.append(f"| {i} | `{r['id']}` | {r['category']} | {result} | {', '.join(bad) if bad else '-'} |")
    lines.append("")

    if failed:
        lines.append("## Failures")
        lines.append("")
        for r in failed:
            lines.append(f"### `{r['id']}` ({r['category']})")
            lines.append("")
            lines.append(f"Request: {r['request']}")
            lines.append("")
            for s in r["steps"]:
                if not s["failures"]:
                    continue
                lines.append(f"Step {s['number']}: `{s['command']}`")
                lines.append("")
                for f in s["failures"]:
                    lines.append(f"- **{f['what']}**: expected `{json.dumps(f['expected'], ensure_ascii=False, default=str)}`, "
                                 f"got `{json.dumps(f['actual'], ensure_ascii=False, default=str)}`")
                diffs = [f["diff"] for f in s["failures"] if f["diff"]]
                if diffs:
                    lines.append("")
                    lines.append("```diff")
                    lines.append("\n\n".join(diffs))
                    lines.append("```")
                if s["stderr"] and s["exit_actual"] != s["exit_expected"]:
                    lines.append("")
                    lines.append("stderr:")
                    lines.append("```")
                    lines.append(s["stderr"])
                    lines.append("```")
                lines.append("")

    if skipped:
        lines.append("## Not runnable (excluded from the score)")
        lines.append("")
        for t in skipped:
            lines.append(f"- `{t['id']}`: {t['not_runnable_reason']}")
        lines.append("")

    return "\n".join(lines) + "\n"


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--python", default=str(DEFAULT_PYTHON), help="venv python to run the CLI with")
    parser.add_argument("--cli", default=str(DEFAULT_CLI), help="path to xlsx_tool.py")
    parser.add_argument("--only", nargs="*", default=None, help="run only these task ids")
    args = parser.parse_args()

    if not TASKS_JSON.exists():
        sys.exit("tasks.json not found: run make_tasks.py first")
    tasks = json.loads(TASKS_JSON.read_text(encoding="utf-8"))
    if args.only:
        tasks = [t for t in tasks if t["id"] in set(args.only)]

    started = dt.datetime.now()
    skipped, results = [], []
    for t in tasks:
        if not t["runnable"]:
            skipped.append(t)
            continue
        print(f"running {t['id']} ...", flush=True)
        results.append(run_task(t, args.python, args.cli))

    RESULTS_MD.write_text(render(results, skipped, args.python, args.cli, started), encoding="utf-8")
    passed = sum(1 for r in results if r["passed"])
    print(f"score: {passed}/{len(results)} runnable passed; results written to {RESULTS_MD}")


if __name__ == "__main__":
    main()
