"""Run the DOCX v2 accuracy benchmark.

Builds each task's input folder, runs the CLI in that folder, compares the
result with the expected result in make_tasks.py, and writes results.md
next to this file.

    python run_benchmark.py                 # all tasks
    python run_benchmark.py --only P1 T4    # some tasks
"""

import argparse
import difflib
import hashlib
import json
import platform
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import make_tasks as mt  # noqa: E402

TOOL = HERE.parent / "scripts" / "docx_tool.py"
DEFAULT_PYTHON = sys.executable
CATEGORY_ORDER = ["read", "create", "replace", "fill", "track-changes", "comments", "validate",
                  "corrupt-file", "summary", "to-pdf", "out-of-scope", "missing-file", "safety"]
MISSING = "<missing>"


# ---------------------------------------------------------------- environment

def libreoffice_present():
    if shutil.which("soffice") or shutil.which("soffice.exe"):
        return True
    return Path(r"C:\Program Files\LibreOffice\program\soffice.exe").exists()


def word_present(python):
    """Word is registered and the benchmark's Python can import pywin32 (the tool needs both)."""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            pass
    except OSError:
        return False
    probe = subprocess.run([python, "-c", "import win32com.client"], capture_output=True, check=False)
    return probe.returncode == 0


def converter_for(python):
    if word_present(python):
        return "word"
    return "libreoffice" if libreoffice_present() else "none"


def snapshot(folder):
    out = {}
    for p in sorted(folder.rglob("*")):
        if p.is_file():
            out[p.relative_to(folder).as_posix()] = hashlib.sha256(p.read_bytes()).hexdigest()
    return out


# ---------------------------------------------------------------- comparison helpers

def strict_eq(a, b):
    """Equality that does not treat True == 1 or False == 0."""
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return a == b
    if isinstance(a, dict) and isinstance(b, dict):
        return a.keys() == b.keys() and all(strict_eq(a[k], b[k]) for k in a)
    if isinstance(a, list) and isinstance(b, list):
        return len(a) == len(b) and all(strict_eq(x, y) for x, y in zip(a, b))
    return type(a) is type(b) and a == b


def resolve(data, path):
    """Dotted path lookup. 'name.0.kind' indexes lists, 'name.#' is the length of a list."""
    cur = data
    for part in path.split("."):
        try:
            if part == "#":
                cur = len(cur)
            elif isinstance(cur, list):
                cur = cur[int(part)]
            else:
                cur = cur[part]
        except (KeyError, IndexError, ValueError, TypeError):
            return MISSING
    return cur


def add_mismatch(failures, name, expected, actual):
    e = json.dumps(expected, indent=2, ensure_ascii=False, sort_keys=True).splitlines()
    a = json.dumps(actual, indent=2, ensure_ascii=False, sort_keys=True).splitlines()
    diff = "\n".join(difflib.unified_diff(e, a, "expected", "actual", lineterm="", n=3))
    failures.append({"check": name, "diff": diff or f"expected: {e}\nactual:   {a}"})


def add_text_mismatch(failures, name, expected, actual):
    e = expected.splitlines() or [""]
    a = actual.splitlines() or [""]
    diff = "\n".join(difflib.unified_diff(e, a, "expected", "actual", lineterm="", n=3))
    failures.append({"check": name, "diff": diff or "(texts differ only in whitespace)"})


# ---------------------------------------------------------------- checks

def check_json(want, data, failures):
    actual = {path: resolve(data, path) for path in want}
    bad = [p for p in want if not strict_eq(want[p], actual[p])]
    if bad:
        add_mismatch(failures, "JSON fields",
                     {p: want[p] for p in bad}, {p: actual[p] for p in bad})


def check_output(spec, folder, failures):
    path = folder / spec["file"]
    if not path.exists():
        add_text_mismatch(failures, "output file exists", spec["file"], "missing")
        return
    if path.suffix.lower() == ".pdf":
        if "pdf_pages" in spec:
            try:
                pages = mt.pdf_pages(path)
            except Exception as exc:  # noqa: BLE001
                add_text_mismatch(failures, "PDF opens", "a readable PDF", repr(exc))
                return
            if pages != spec["pdf_pages"]:
                add_mismatch(failures, "PDF page count", spec["pdf_pages"], pages)
        return
    try:
        f = mt.collect_facts(path)
    except Exception as exc:  # noqa: BLE001
        add_text_mismatch(failures, "output opens as .docx", "a readable .docx", repr(exc))
        return
    for key, want in spec.items():
        if key == "file":
            continue
        if key == "body_paragraphs":
            if not strict_eq(want, f["body"]):
                add_mismatch(failures, "body paragraphs (style, text)", want, f["body"])
        elif key == "contains":
            missing = [s for s in want if s not in f["visible"]]
            if missing:
                add_mismatch(failures, "text present in file", want,
                             [s for s in want if s in f["visible"]])
        elif key == "absent":
            present = [s for s in want if s in f["visible"]]
            if present:
                add_mismatch(failures, "text absent from file", [],
                             {"still_present": present})
        elif key == "occurrences":
            actual = {s: f["visible"].count(s) for s in want}
            if not strict_eq(want, actual):
                add_mismatch(failures, "occurrence counts in file", want, actual)
        elif key == "revisions":
            actual = {k: f["revisions"][k] for k in want}
            if not strict_eq(want, actual):
                add_mismatch(failures, "tracked-change counts in XML", want, actual)
        elif key == "deleted_contains":
            missing = [s for s in want if s not in f["deleted"]]
            if missing:
                add_mismatch(failures, "deleted text present", want,
                             [s for s in want if s in f["deleted"]])
        elif key == "comments":
            subset = [{k: a.get(k) for k in w} for w, a in zip(want, f["comments"])]
            if len(want) != len(f["comments"]) or not strict_eq(want, subset):
                shown = subset if len(want) == len(f["comments"]) else f["comments"]
                add_mismatch(failures, "comments in file", want, shown)
        elif key == "run_format":
            matches = [r for r in f["runs"] if want["text"] in r["text"]]
            ok = bool(matches) and all(r["bold"] == want["bold"] and r["italic"] == want["italic"]
                                       for r in matches)
            if not ok:
                add_mismatch(failures, f"formatting of run containing {want['text']!r}",
                             {"bold": want["bold"], "italic": want["italic"]},
                             [{"text": r["text"], "bold": r["bold"], "italic": r["italic"]}
                              for r in matches] or "no run contains the text")
        else:
            failures.append({"check": "unknown expectation key", "diff": key})


def evaluate(expect, rc, out, err, before, after, folder, failures):
    if rc != expect["exit"]:
        add_mismatch(failures, "exit code", expect["exit"], rc)

    combined = out + "\n" + err
    if "message" in expect:
        missing = [m for m in expect["message"] if m not in combined]
        if missing:
            add_mismatch(failures, "error/output message contains", expect["message"],
                         [m for m in expect["message"] if m in combined])

    if expect.get("stderr_nonempty") and not err.strip():
        add_text_mismatch(failures, "stderr has an error line", "non-empty stderr", "(empty)")

    if "stdout_equals" in expect:
        got = out.replace("\r\n", "\n").rstrip("\n")
        if got != expect["stdout_equals"]:
            add_text_mismatch(failures, "stdout text", expect["stdout_equals"], got)

    if "json" in expect or expect.get("checks_all_passed"):
        try:
            data = json.loads(out)
        except ValueError:
            add_text_mismatch(failures, "stdout is one JSON object", "JSON object", out[:300])
            data = None
        if data is not None and "json" in expect:
            check_json(expect["json"], data, failures)
        if data is not None and expect.get("checks_all_passed"):
            checks = data.get("checks", [])
            bad = [c.get("name") for c in checks if c.get("passed") is not True]
            if not checks or bad:
                add_mismatch(failures, "every validation check passes", {"failed_checks": []},
                             {"failed_checks": bad or ["no checks listed"]})

    created = sorted(set(after) - set(before))
    if expect.get("new_files") is not None:
        want = sorted(expect["new_files"])
        if want != created:
            add_mismatch(failures, "files written next to input", want, created)

    changed = sorted(n for n in before if n not in after or after[n] != before[n])
    want_changed = sorted(expect.get("changed", []))
    if want_changed != changed:
        add_mismatch(failures, "pre-existing files changed", want_changed, changed)

    if "output" in expect:
        check_output(expect["output"], folder, failures)


def run_task(task, work, args, converter):
    res = {"id": task["id"], "category": task["category"], "title": task["title"],
           "request": task["request"], "argv": task["argv"], "status": "PASS",
           "failures": [], "rc": None, "want_exit": None, "stdout": "", "stderr": "",
           "note": ""}
    folder = work / task["id"]
    folder.mkdir(parents=True)
    try:
        if task["fixture"] is not None:
            task["fixture"](folder)
    except Exception as exc:  # noqa: BLE001
        res["status"] = "NOT RUNNABLE"
        res["note"] = f"benchmark fixture failed: {exc!r}"
        return res

    expect = task["expect"]
    if "variants" in expect:
        expect = expect["variants"][converter]
        res["note"] = f"to-pdf converter: {converter} (variant used)"
    res["want_exit"] = expect.get("exit")

    before = snapshot(folder)
    argv = [args.python, str(TOOL)] + [a.replace("{dir}", str(folder)) for a in task["argv"]]
    res["shown"] = " ".join(["docx_tool.py"] + [a.replace("{dir}", "<task folder>")
                                                for a in task["argv"]])
    try:
        proc = subprocess.run(argv, cwd=str(folder), capture_output=True,
                              timeout=args.timeout, stdin=subprocess.DEVNULL)
    except subprocess.TimeoutExpired:
        res["status"] = "FAIL"
        res["failures"].append({"check": "finished in time",
                                "diff": f"timed out after {args.timeout} seconds"})
        return res

    rc = proc.returncode
    out = proc.stdout.decode("utf-8", "replace")
    err = proc.stderr.decode("utf-8", "replace")
    res.update(rc=rc, stdout=out, stderr=err)
    after = snapshot(folder)
    try:
        evaluate(expect, rc, out, err, before, after, folder, res["failures"])
    except Exception as exc:  # noqa: BLE001
        res["failures"].append({"check": "benchmark runner error", "diff": repr(exc)})
    res["status"] = "PASS" if not res["failures"] else "FAIL"
    return res


# ---------------------------------------------------------------- report

def _cell(text):
    return str(text).replace("|", "\\|")


def write_results(results, meta, path):
    runnable = [r for r in results if r["status"] != "NOT RUNNABLE"]
    passed = [r for r in runnable if r["status"] == "PASS"]
    failed = [r for r in runnable if r["status"] == "FAIL"]
    lines = ["# DOCX v2 accuracy benchmark results", ""]
    lines += [f"- Run: {meta['when']}",
              f"- CLI under test: `{meta['tool']}`",
              f"- Python: `{meta['python']}` ({meta['platform']})",
              f"- Word found (with pywin32): {meta['word']}",
              f"- LibreOffice found: {meta['lo']}",
              f"- to-pdf variant used: {meta['converter']}",
              "- Expected results: written by hand from `specs/docx.md` and `docx/SKILL.md`, "
              "never taken from a run of the tool.", ""]
    lines += [f"## Score: {len(passed)}/{len(runnable)} passed", ""]
    not_run = [r for r in results if r["status"] == "NOT RUNNABLE"]
    if not_run:
        lines += ["Not runnable (benchmark setup error, not counted):"]
        lines += [f"- {r['id']}: {r['note']}" for r in not_run]
        lines.append("")

    lines += ["## Score by category", "", "| Category | Passed | Runnable |", "|---|---|---|"]
    cats = [c for c in CATEGORY_ORDER if any(r["category"] == c for r in results)]
    cats += sorted({r["category"] for r in results} - set(cats))
    for c in cats:
        rs = [r for r in runnable if r["category"] == c]
        lines.append(f"| {c} | {sum(r['status'] == 'PASS' for r in rs)} | {len(rs)} |")
    lines.append("")

    lines += ["## Top failure categories", ""]
    counts = {}
    for r in failed:
        counts[r["category"]] = counts.get(r["category"], 0) + 1
    if counts:
        for c, n in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
            ids = ", ".join(r["id"] for r in failed if r["category"] == c)
            lines.append(f"- {c}: {n} failed ({ids})")
    else:
        lines.append("None.")
    lines.append("")

    lines += ["## Per-task results", "",
              "| ID | Category | Task | Result | Exit (got / want) |", "|---|---|---|---|---|"]
    for r in results:
        got = "-" if r["rc"] is None else r["rc"]
        want = "-" if r["want_exit"] is None else r["want_exit"]
        lines.append(f"| {r['id']} | {r['category']} | {_cell(r['title'])} | {r['status']} | "
                     f"{got} / {want} |")
    lines.append("")

    if failed or not_run:
        lines += ["## Failures", ""]
        for r in failed:
            lines += [f"### {r['id']}: {r['title']} (FAIL)", "",
                      f"- Request: {r['request']}",
                      f"- Command: `{r.get('shown', '')}`"]
            if r["note"]:
                lines.append(f"- Note: {r['note']}")
            lines.append("")
            for f in r["failures"]:
                lines += [f"**{f['check']}**", "", "```diff", f["diff"], "```", ""]
            lines += ["stdout (first 600 characters):", "", "```",
                      r["stdout"][:600].rstrip() or "(empty)", "```", "",
                      "stderr (first 600 characters):", "", "```",
                      r["stderr"][:600].rstrip() or "(empty)", "```", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description="Run the DOCX v2 accuracy benchmark.")
    ap.add_argument("--python", default=DEFAULT_PYTHON, help="Python in the benchmark venv")
    ap.add_argument("--work", default=str(HERE / "work"), help="scratch folder for task inputs")
    ap.add_argument("--out", default=str(HERE / "results.md"), help="results file")
    ap.add_argument("--timeout", type=int, default=180, help="seconds per CLI call")
    ap.add_argument("--only", nargs="*", help="run only these task ids")
    args = ap.parse_args()

    if not TOOL.exists():
        sys.exit(f"CLI not found: {TOOL}")
    work = Path(args.work)
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)

    converter = converter_for(args.python)
    lo = libreoffice_present()
    tasks = mt.build_tasks()
    if args.only:
        tasks = [t for t in tasks if t["id"] in set(args.only)]

    results = []
    for task in tasks:
        res = run_task(task, work, args, converter)
        results.append(res)
        print(f"{res['id']:<5} {res['status']}")

    meta = {"when": datetime.now().strftime("%Y-%m-%d %H:%M"), "tool": str(TOOL),
            "python": args.python, "platform": platform.platform(),
            "lo": "yes" if lo else "no", "word": "yes" if converter == "word" else "no",
            "converter": converter}
    write_results(results, meta, Path(args.out))

    runnable = [r for r in results if r["status"] != "NOT RUNNABLE"]
    passed = sum(r["status"] == "PASS" for r in runnable)
    print(f"Score: {passed}/{len(runnable)} passed. Results: {args.out}")


if __name__ == "__main__":
    main()
