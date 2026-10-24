#!/usr/bin/env python3
"""Run the PDF skill v2 benchmark.

1. Rebuilds fixtures and tasks.json with make_tasks.py (expected values live there).
2. Runs each task through the CLI from a fresh work folder under runs/.
3. Compares exit code, JSON, stdout/stderr, and output PDFs (read with pypdf, not with the tool).
4. Writes results.md with per-task PASS/FAIL, diffs for failures, and scores by category.

Usage:  <venv python> run_benchmark.py [--only SUBSTRING] [--timeout 120]
"""
import argparse
import datetime
import difflib
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
FIX = HERE / "fixtures"
RUNS = HERE / "runs"
MANIFEST = HERE / "tasks.json"
RESULTS = HERE / "results.md"
DEFAULT_TOOL = Path(__file__).resolve().parents[1] / "scripts" / "pdf_tool.py"
DEFAULT_PY = Path(sys.executable)
SNIP = 1500
PASSWORD_ENV_NAMES = ["PDF_TEST_PW", "PDF_PASSWORD"]


# ------------------------------------------------------------ availability

NEEDS = {
    "tesseract": lambda: shutil.which("tesseract") is not None,
    "poppler": lambda: shutil.which("pdftoppm") is not None or shutil.which("pdftotext") is not None,
}


# ------------------------------------------------------------ helpers

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def snapshot(d):
    if not d.exists():
        return {}
    return {str(p.relative_to(d)).replace("\\", "/"): sha(p) for p in d.rglob("*") if p.is_file()}


def snip(s, n=SNIP):
    s = s or ""
    return s if len(s) <= n else s[:n] + f"\n... [truncated, {len(s)} chars total]"


def dumps(v):
    return json.dumps(v, indent=2, ensure_ascii=False)


def strict_eq(a, b):
    """Equality where True is not 1 and 1 is not True."""
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if isinstance(a, dict) and isinstance(b, dict):
        return a.keys() == b.keys() and all(strict_eq(a[k], b[k]) for k in a)
    if isinstance(a, list) and isinstance(b, list):
        return len(a) == len(b) and all(strict_eq(x, y) for x, y in zip(a, b))
    return a == b


def resolve(obj, path):
    """Dotted lookup. 'len' returns the length of the current list or dict."""
    cur = obj
    for seg in path.split("."):
        if seg == "len":
            cur = len(cur)
        elif isinstance(cur, list):
            cur = cur[int(seg)]
        else:
            cur = cur[seg]
    return cur


def norm(s):
    return " ".join((s or "").split())


def _val(x):
    return x.get_object() if hasattr(x, "get_object") else x


def field_value(f):
    try:
        if "/V" in f:
            return _val(f["/V"])
    except Exception:
        pass
    return getattr(f, "value", None)


def bool_of(v):
    if isinstance(v, bool):
        return v
    s = "" if v is None else str(v)
    if s in ("/Yes", "/On", "Yes", "On", "true", "True"):
        return True
    if s in ("/Off", "Off", "false", "False", ""):
        return False
    return None


def pdf_facts(path):
    from pypdf import PdfReader
    r = PdfReader(str(path))
    return {
        "page_count": len(r.pages),
        "texts": [norm(p.extract_text()) for p in r.pages],
        "rots": [int(_val(p.get("/Rotate", 0))) % 360 for p in r.pages],
        "fields": r.get_fields() or {},
    }


def page_at(lst, pg):
    i = int(pg) - 1
    return lst[i] if 0 <= i < len(lst) else None


# ------------------------------------------------------------ checks

def evaluate(task, ctx):
    recs = []

    def rec(label, ok, exp="", act=""):
        recs.append({"label": label, "ok": bool(ok), "expected": str(exp), "actual": str(act)})

    rec("exit code", ctx["exit"] == task["exit"], task["exit"], ctx["exit"])
    js = ctx["json"]
    work = ctx["work"]

    for c in task["checks"]:
        t = c["t"]

        if t == "json_equals":
            for path, exp in c["map"].items():
                label = f"json {path} == {json.dumps(exp, ensure_ascii=False)}"
                if js is None:
                    rec(label, False, dumps(exp), "stdout is not a JSON object")
                    continue
                try:
                    act = resolve(js, path)
                except Exception as e:
                    rec(label, False, dumps(exp), f"(path not found: {type(e).__name__})")
                    continue
                rec(label, strict_eq(act, exp), dumps(exp), dumps(act))

        elif t == "json_contains":
            for path, subs in c["map"].items():
                if js is None:
                    rec(f"json {path} contains", False, subs, "stdout is not a JSON object")
                    continue
                try:
                    text = resolve(js, path)
                except Exception as e:
                    rec(f"json {path} contains", False, subs, f"(path not found: {type(e).__name__})")
                    continue
                text = text if isinstance(text, str) else dumps(text)
                for s in subs:
                    rec(f"json {path} contains {s!r}", s in text, f"contains {s}", snip(text, 600))

        elif t == "json_len_min":
            if js is None:
                rec(f"json {c['path']} length >= {c['min']}", False, f">= {c['min']}", "stdout is not JSON")
                continue
            try:
                n = len(resolve(js, c["path"]))
            except Exception as e:
                rec(f"json {c['path']} length >= {c['min']}", False, f">= {c['min']}", f"(path not found: {e})")
                continue
            rec(f"json {c['path']} length >= {c['min']}", n >= c["min"], f">= {c['min']}", n)

        elif t == "json_items":
            label = f"json {c['path']} items keyed by {c['key']}"
            if js is None:
                rec(label, False, dumps(c["expect"]), "stdout is not a JSON object")
                continue
            try:
                items = resolve(js, c["path"])
            except Exception as e:
                rec(label, False, dumps(c["expect"]), f"(path not found: {e})")
                continue
            for e in c["expect"]:
                kv = e[c["key"]]
                match = next((a for a in items if isinstance(a, dict) and a.get(c["key"]) == kv), None)
                sub_label = f"  item {kv}"
                if match is None:
                    rec(sub_label, False, dumps(e), "no item with this key")
                    continue
                actual_sub = {k: match.get(k, "<missing>") for k in e}
                rec(sub_label, strict_eq(actual_sub, e), dumps(e), dumps(actual_sub))

        elif t == "stdout_empty":
            rec("stdout is empty", ctx["stdout"].strip() == "", "(empty)", snip(ctx["stdout"], 300))

        elif t == "stderr_contains":
            for s in c["items"]:
                rec(f"stderr contains {s!r} (case-insensitive)", s.lower() in ctx["stderr"].lower(),
                    f"contains {s}", snip(ctx["stderr"], 600))

        elif t == "no_secret":
            for v in c["values"]:
                leaked = v in ctx["stdout"] or v in ctx["stderr"]
                rec(f"password value not printed ({len(v)} chars)", not leaked, "absent", "PRESENT in output"
                    if leaked else "absent")

        elif t == "file_exists":
            exists = (work / c["path"]).exists()
            rec(f"file {c['path']} exists == {c['exists']}", exists == c["exists"], c["exists"], exists)

        elif t == "no_pdf_in":
            d = work / c["path"]
            ok = (not d.exists()) or not any(d.glob("*.pdf"))
            found = sorted(p.name for p in d.glob("*.pdf")) if d.exists() else []
            rec(f"no PDF written in {c['path']}/", ok, "no .pdf files", found or "none")

        elif t == "file_unchanged":
            pre = ctx["pre"].get(c["path"])
            post = ctx["post"].get(c["path"])
            rec(f"input {c['path']} unchanged", pre is not None and pre == post, "same sha256 as before run",
                "changed or removed" if pre != post else "same")

        elif t == "pdf":
            p = work / c["path"]
            label = f"pdf {c['path']}"
            if not p.exists():
                rec(f"{label} exists", False, "file", "missing")
                continue
            try:
                f = pdf_facts(p)
            except Exception as e:
                rec(f"{label} opens with pypdf", False, "readable PDF", f"{type(e).__name__}: {e}")
                continue
            if "page_count" in c:
                rec(f"{label} page count", f["page_count"] == c["page_count"], c["page_count"], f["page_count"])
            for pg, want in c.get("rotations", {}).items():
                got = page_at(f["rots"], pg)
                rec(f"{label} page {pg} /Rotate", got == want, want, got)
            for pg, subs in c.get("page_text_contains", {}).items():
                txt = page_at(f["texts"], pg) or ""
                for s in subs:
                    rec(f"{label} page {pg} text contains {s!r}", s in txt, f"contains {s}", snip(txt, 600))
            for pg, seq in c.get("page_text_sequence", {}).items():
                txt = page_at(f["texts"], pg) or ""
                pos = [txt.find(s) for s in seq]
                ok = all(x >= 0 for x in pos) and pos == sorted(pos) and len(set(pos)) == len(pos)
                rec(f"{label} page {pg} has {seq} in order", ok, " then ".join(seq), snip(txt, 600))
            for name, want in c.get("fields", {}).items():
                fields = f["fields"]
                if name not in fields:
                    rec(f"{label} field {name}", False, repr(want), "field not present")
                    continue
                raw = field_value(fields[name])
                if isinstance(want, bool):
                    got = bool_of(raw)
                    rec(f"{label} field {name}", got == want, want, f"{got} (raw {raw!r})")
                else:
                    got = "" if raw is None else str(raw)
                    rec(f"{label} field {name}", got == want, repr(want), repr(got))

        else:
            rec(f"unknown check type {t}", False, "", "")

    return recs


# ------------------------------------------------------------ run

def run_one(task, py, tool, timeout):
    work = RUNS / task["id"]
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    for wname, fname in task.get("copy", {}).items():
        dst = work / wname
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(FIX / fname, dst)
    for wname, text in task.get("write", {}).items():
        dst = work / wname
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_text(text, encoding="utf-8")

    pre = snapshot(work)
    env = os.environ.copy()
    for name in PASSWORD_ENV_NAMES:
        env.pop(name, None)
    env.update(task.get("env", {}))
    cmd = [str(py), str(tool), *task["args"]]

    t0 = time.time()
    try:
        proc = subprocess.run(cmd, cwd=str(work), env=env, capture_output=True, timeout=timeout)
        code = proc.returncode
        out = proc.stdout.decode("utf-8", "replace")
        err = proc.stderr.decode("utf-8", "replace")
        error = None
    except subprocess.TimeoutExpired:
        code, out, err, error = None, "", "", f"timeout after {timeout} s"
    elapsed = time.time() - t0

    js = None
    if out.strip():
        try:
            js = json.loads(out)
        except Exception:
            js = None

    post = snapshot(work)
    ctx = {"work": work, "exit": code if code is not None else -999, "stdout": out, "stderr": err,
           "json": js, "pre": pre, "post": post}
    recs = evaluate(task, ctx) if error is None else [
        {"label": "exit code", "ok": False, "expected": str(task["exit"]), "actual": error}]
    passed = all(r["ok"] for r in recs)
    return {"id": task["id"], "category": task["category"], "kind": task["kind"], "title": task["request"],
            "cmd": " ".join(["python", "pdf_tool.py", *task["args"]]), "exit_expected": task["exit"],
            "exit_actual": code, "stdout": out, "stderr": err, "checks": recs, "passed": passed,
            "elapsed": round(elapsed, 2), "error": error}


def build_report(results, not_runnable, tool, py, started):
    runnable = [r for r in results]
    total_pass = sum(1 for r in runnable if r["passed"])
    cats = []
    kinds = []
    for r in results:
        if r["category"] not in cats:
            cats.append(r["category"])
        if r["kind"] not in kinds:
            kinds.append(r["kind"])

    def score_table(key, names):
        lines = ["| Group | Passed | Runnable | Score |", "|---|---|---|---|"]
        for n in names:
            rows = [r for r in runnable if r[key] == n]
            p = sum(1 for r in rows if r["passed"])
            pct = f"{100 * p // len(rows)}%" if rows else "n/a"
            lines.append(f"| {n} | {p} | {len(rows)} | {pct} |")
        return "\n".join(lines)

    out = []
    out.append("# PDF skill v2 benchmark results\n")
    out.append(f"- Run: {started.strftime('%Y-%m-%d %H:%M')} (local)")
    out.append(f"- Python: {sys.version.split()[0]} at `{py}`")
    out.append(f"- Platform: {platform.platform()}")
    out.append(f"- Tool under test: `{tool}`")
    out.append(f"- Tasks: {len(results) + len(not_runnable)} | Runnable: {len(runnable)} | "
               f"Not runnable: {len(not_runnable)}")
    pct = f"{100 * total_pass // len(runnable)}%" if runnable else "n/a"
    out.append(f"- **Overall score: {total_pass}/{len(runnable)} ({pct})**\n")

    out.append("## Score by category\n")
    out.append(score_table("category", cats) + "\n")
    out.append("## Score by task kind (happy path, refusal, safety)\n")
    out.append(score_table("kind", kinds) + "\n")

    out.append("## Per-task results\n")
    out.append("| ID | Category | Kind | Result | Failed checks |")
    out.append("|---|---|---|---|---|")
    for r in results:
        failed = [c["label"] for c in r["checks"] if not c["ok"]]
        res = "PASS" if r["passed"] else "FAIL"
        out.append(f"| {r['id']} | {r['category']} | {r['kind']} | {res} | "
                   f"{len(failed)} of {len(r['checks'])} checks failed |")
    for nr in not_runnable:
        out.append(f"| {nr['id']} | {nr['category']} | {nr['kind']} | NOT RUNNABLE | {nr['reason']} |")
    out.append("")

    if not_runnable:
        out.append("## Not runnable\n")
        for nr in not_runnable:
            out.append(f"- {nr['id']}: {nr['reason']}")
        out.append("")

    fails = [r for r in results if not r["passed"]]
    out.append("## Failures\n")
    if not fails:
        out.append("No failures.\n")
    for r in fails:
        out.append(f"### {r['id']} ({r['category']}, {r['kind']})\n")
        out.append(f"- Request: {r['title']}")
        out.append(f"- Command: `{r['cmd']}`")
        out.append(f"- Exit: expected {r['exit_expected']}, actual {r['exit_actual']}")
        if r["error"]:
            out.append(f"- Error: {r['error']}")
        out.append(f"- Elapsed: {r['elapsed']} s\n")
        out.append("stdout:\n```\n" + snip(r["stdout"], SNIP).rstrip() + "\n```\n")
        out.append("stderr:\n```\n" + snip(r["stderr"], SNIP).rstrip() + "\n```\n")
        for i, c in enumerate([c for c in r["checks"] if not c["ok"]], 1):
            out.append(f"{i}. **{c['label']}** - failed")
            diff = "\n".join(difflib.unified_diff(
                c["expected"].splitlines() or [""], c["actual"].splitlines() or [""],
                "expected", "actual", lineterm=""))
            out.append("```diff\n" + diff + "\n```\n")
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--python", default=str(DEFAULT_PY))
    ap.add_argument("--tool", default=str(DEFAULT_TOOL))
    ap.add_argument("--only", default=None, help="run only task ids containing this text")
    ap.add_argument("--timeout", type=int, default=120)
    ap.add_argument("--no-rebuild", action="store_true", help="skip make_tasks.py")
    a = ap.parse_args()

    py, tool = Path(a.python), Path(a.tool)
    started = datetime.datetime.now()
    if not a.no_rebuild or not MANIFEST.exists():
        subprocess.run([str(py), str(HERE / "make_tasks.py")], check=True)
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))

    RUNS.mkdir(exist_ok=True)
    results, not_runnable = [], []
    for task in manifest["tasks"]:
        if a.only and a.only not in task["id"]:
            continue
        missing = [n for n in task.get("needs", []) if not NEEDS[n]()]
        if missing:
            not_runnable.append({"id": task["id"], "category": task["category"], "kind": task["kind"],
                                 "reason": "needs " + ", ".join(missing) + " (not installed)"})
            print(f"NOT RUNNABLE {task['id']}")
            continue
        r = run_one(task, py, tool, a.timeout)
        results.append(r)
        print(f"{'PASS' if r['passed'] else 'FAIL'} {r['id']} ({r['elapsed']} s)")

    RESULTS.write_text(build_report(results, not_runnable, tool, py, started), encoding="utf-8")
    passed = sum(1 for r in results if r["passed"])
    print(f"\nScore: {passed}/{len(results)} runnable. Report: {RESULTS}")


if __name__ == "__main__":
    main()
