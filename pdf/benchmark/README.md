# PDF skill v2 accuracy benchmark

Measures the v2 PDF skill (`pdf_tool.py`, documented in `../SKILL.md`, specified in `../../specs/pdf.md`)
against 30 tasks. Each task is a request a business user might make, run through the real CLI.

## Files

- `make_tasks.py` builds the input PDFs in `fixtures/` (reportlab and pypdf) and writes `tasks.json`,
  the 30 task definitions with their expected results.
- `run_benchmark.py` rebuilds the manifest, runs each task in a fresh folder under `runs/`, checks the
  result, and writes `results.md`.
- `tasks.json`, `fixtures/`, `runs/`, `results.md` are generated. Re-run to refresh them.

## How to run

Use the venv that has pypdf, pdfplumber, reportlab and cryptography installed:

    <your-venv-folder>\venv-v2-pdf\Scripts\python.exe run_benchmark.py

Options: `--only 14_` runs matching task ids, `--timeout 120` sets the per-command limit in seconds,
`--no-rebuild` skips regenerating fixtures. The tool under test is set by `--tool`, which defaults to
`...\office-skills-v2\pdf\scripts\pdf_tool.py`.

Only `make_tasks.py` and `run_benchmark.py` are needed. Neither modifies the skill.

## Categories

Each task has a command group (category) and a kind:

- `happy`: a normal request that should succeed and produce the right output.
- `refusal`: a request the spec says must be refused with a given exit code and message.
- `safety`: corrupt or empty input, password-protected input without or with a wrong password, and
  inputs that must not be overwritten.

| Category | Tasks | What it covers |
|---|---|---|
| text | 01-03 | extraction, page list, scanned PDF refused (no text layer) |
| password | 04-06 | no password refused, correct password via env var, wrong password, password never printed |
| tables | 07 | table rows from a page |
| fields | 08-11 | fields-list, fields-fill with verification, unknown field writes nothing, no form |
| merge | 12-13 | merge two PDFs, existing output refused without --force |
| split | 14-15 | split every N pages, range with an uncovered page |
| rotate | 16-19 | rotate a page, modulo 360, input overwrite refused, missing --angle |
| reorder | 20-21 | reorder by permutation, non-permutation refused |
| watermark | 22-23 | stamp and text watermark, text preserved |
| create | 24-25 | text to PDF, unsupported character refused (out of scope) |
| check | 26-29 | valid, truncated (corrupt), empty file, encrypted without password |
| check-text | 30 | phrases found and missing, case-sensitive matching |

Score = passed / runnable, overall and by category and kind. A task whose tools (Tesseract, Poppler)
are missing is marked NOT RUNNABLE and left out of the denominator. The `needs` field in
`tasks.json` drives this. No task needs Tesseract or Poppler, because OCR is out of scope in the spec
and is tested as a refusal (task 02).

## Expected-result method

- Expected values are written by hand in `make_tasks.py`, from `specs/pdf.md` (sections 2, 3, 4) and
  `SKILL.md`. They are not taken from any tool output. The fixture generator asserts the facts the
  tasks depend on (page texts, field names, encryption, page count, file sizes) using pypdf only.
- Each task lists its checks: exit code (always), JSON values with strict typing (`True` is not `1`),
  JSON item matching by key, stdout or stderr text, password never printed, file exists or absent,
  input unchanged by SHA-256, and output PDFs read independently with pypdf (page count, `/Rotate`,
  page text in order, form field values, normalised whitespace).
- The grader never trusts the tool's own `"verified": true`. It re-reads the output PDF itself.
- Failures show the expected and actual values as a unified diff in `results.md`.
- Harness check: a deliberately wrong expectation (task 16, rotation 180 instead of 90) is reported as
  a failure, so a pass means the checks ran and matched.

## Limits

- 30 tasks is a sample, not full coverage. Visual layout, rendering and tables with no ruling lines
  are not tested.
- Fixtures are generated with reportlab, not with LibreOffice as the spec suggests for the acceptance
  fixtures. Text positions and extraction can differ slightly from the spec's fixtures.
- Watermark text is checked by extraction only, which depends on how the rotated text extracts.
