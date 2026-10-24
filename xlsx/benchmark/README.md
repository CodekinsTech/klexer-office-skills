# XLSX v2 skill accuracy benchmark

This benchmark checks the v2 `xlsx` skill (`xlsx/scripts/xlsx_tool.py`) against 30 tasks a business user might give a bot. Each task is a small workbook, one or more CLI commands, and the result the spec says must come out. A task passes only if every step's exit code and every check match.

## Files

| File | Purpose |
|---|---|
| `make_tasks.py` | Builds the input files for each task (with openpyxl) and writes `tasks.json` with the expected results. Does not run the tool. |
| `run_benchmark.py` | Runs each task's steps through the CLI, compares the results, and writes `results.md`. |
| `README.md` | This file. |
| `tasks.json` | Generated: the 30 task definitions, including expected results. |
| `cases/<task_id>/` | Generated: the input files for each task. |
| `work/<task_id>/` | Generated on each run: a fresh copy of the case folder that the CLI runs in, including its outputs. |
| `results.md` | Generated: per-task PASS/FAIL, diffs for failures, scores by category. |

## How to run

Use the benchmark venv (openpyxl 3.1 is required). From the `xlsx/benchmark` folder:

```
<your-venv-folder>\venv-v2-xlsx\Scripts\python.exe make_tasks.py
<your-venv-folder>\venv-v2-xlsx\Scripts\python.exe run_benchmark.py
```

Options for `run_benchmark.py`:

- `--only TASK_ID [TASK_ID ...]` runs selected tasks, for example `--only errors_div0_after_recalc`.
- `--python PATH` chooses the Python that runs the CLI (default: the venv above).
- `--cli PATH` chooses the CLI file (default: `../scripts/xlsx_tool.py`).

`make_tasks.py` rebuilds `cases/` and `tasks.json` from scratch each time. Re-run it after changing a task. The benchmark does not change the skill; it only reads and runs it.

## Score

`Score = passed / runnable`. A task is runnable when its inputs could be built. A task whose setup fails is listed under "Not runnable" in `results.md` and is left out of the score. Setup failures are reported, never counted as passes.

## Categories

| Category | Tasks | What it covers |
|---|---:|---|
| sheets | 1 | Listing sheets, with visible and hidden states and dimensions. |
| read | 3 | Values and formula text, cached values from an Excel-saved file, the "not calculated" flag. |
| write | 5 | Formulas with the `_xlfn.` prefix, dates, value types and clearing a cell, and rejected patches (formula without `=`, invalid date). |
| format | 2 | The currency preset, and refusal when both `--preset` and `--code` are given. |
| recalc | 8 | Arithmetic and operators, VLOOKUP, INDEX/MATCH and XLOOKUP, SUMIFS family, LET and IFERROR, FILTER/SORT/UNIQUE/SEQUENCE spills, defined names, cross-sheet references with TEXTJOIN and MAXIFS/MINIFS, and a blocked spill that shows `#SPILL!`. |
| refusal | 5 | Missing and unsupported files (missing, legacy `.xls`, CSV, macro-enabled `.xlsm`), LAMBDA and unsorted approximate lookups refused with nothing written, password-protected file (not runnable in this venv, see decision 5), and the drawing guard with `--allow-drawing-rewrite`. |
| errors | 2 | Formula errors found after a recalc, and formulas with no calculated value. |
| validate | 1 | A valid workbook, and a truncated copy that must be reported as invalid. |
| check-values | 3 | The rules-file example from the spec, unique/allowed/max_length rules, and a rules file that names a missing column. |

## Expected-result method

- **Source of truth:** `specs/xlsx.md` and `SKILL.md` only. Each expected value is typed into `make_tasks.py` by hand, with the arithmetic in the task definition (for example, `B2 = A1*2 + A2/4 = 20 + 5 = 25`). The tool is never run to make an expected value.
- **Inputs:** workbooks built with openpyxl. Formulas in input files are stored the way Excel saves them (for example `_xlfn._xlws.FILTER`, `_xlfn.XLOOKUP`, `_xlfn._xlpm.x` for LET names), because that is what real files contain.
- **Cached values:** where a task needs a stored result (for example, a file saved by Excel), `make_tasks.py` writes the cached `<v>` into the sheet XML and reads it back with openpyxl before continuing.
- **Output checks:** output workbooks are read back with openpyxl, not with the tool. Values, formula text and number formats are compared. Dates are compared as ISO strings. Booleans must be booleans. Numbers match within 1e-9.
- **JSON checks:** parsed stdout must match the expected value. Dictionaries must have exactly the listed keys. Lists marked `unordered` are sorted before comparison.
- **Messages:** stderr must contain the exact message text from the spec's error table (with the cell and sheet filled in).
- **Exit codes:** each step has its expected exit code from the spec.

## Interpretation decisions

Where the spec is open, the benchmark makes a stated choice:

1. **Duplicate values (`unique`):** the first occurrence passes and each later occurrence is a violation. The spec says "repeated values fail", which could also mean both occurrences fail.
2. **XLOOKUP `if_not_found`:** used in one task (`recalc_lookup_functions`). It is standard XLOOKUP, but the spec does not list it.
3. **Spills:** checked through the cell values in the output file. The `computed` JSON is checked only for single-value formulas.
4. **LAMBDA:** the test uses the immediately invoked form `LAMBDA(x, x*2)(5)`, which is how a user would write it. The spec requires the message to name LAMBDA.
5. **Password-protected file:** the case needs `msoffcrypto-tool` to encrypt a workbook. That package is not in the benchmark venv, so the task is listed as not runnable. Installing it into the venv makes the task runnable.
6. **Test-file construction:** `legacy.xls` is an OLE header stub, `macro.xlsm` is an xlsx package with the macro-enabled content type, and `broken.xlsx` is the first 1000 bytes of a valid workbook. Each is a stand-in for the real file type.
