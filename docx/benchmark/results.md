# DOCX v2 accuracy benchmark results

- Run: 2026-10-25 02:58
- CLI under test: `<skill-folder>\docx\scripts\docx_tool.py`
- Python: `<your-venv-folder>\venv-v2-docx\Scripts\python.exe` (Windows-10-10.0.26200-SP0)
- Word found (with pywin32): no
- LibreOffice found: no
- to-pdf variant used: none
- Expected results: written by hand from `specs/docx.md` and `docx/SKILL.md`, never taken from a run of the tool.

## Score: 30/30 passed

## Score by category

| Category | Passed | Runnable |
|---|---|---|
| read | 1 | 1 |
| create | 3 | 3 |
| replace | 4 | 4 |
| fill | 4 | 4 |
| track-changes | 4 | 4 |
| comments | 3 | 3 |
| validate | 1 | 1 |
| corrupt-file | 1 | 1 |
| summary | 1 | 1 |
| to-pdf | 1 | 1 |
| out-of-scope | 4 | 4 |
| missing-file | 1 | 1 |
| safety | 2 | 2 |

## Top failure categories

None.

## Per-task results

| ID | Category | Task | Result | Exit (got / want) |
|---|---|---|---|---|
| R1 | read | Read paragraphs and tables in order | PASS | 0 / 0 |
| C1 | create | Outline with title, body and bullets | PASS | 0 / 0 |
| C2 | create | Heading 2, numbered list, UTF-8 text | PASS | 0 / 0 |
| C3 | create | Outline with no text lines | PASS | 1 / 1 |
| P1 | replace | Replace in body, table cell and header (default output name) | PASS | 0 / 0 |
| P2 | replace | Match split across runs keeps first run's formatting | PASS | 0 / 0 |
| P3 | replace | Pairs file applied in order, with a zero-match pair | PASS | 0 / 0 |
| P4 | replace | Skip matches inside tracked insertions, ignore deletions | PASS | 0 / 0 |
| F1 | fill | Fill placeholders in body, table and footer | PASS | 0 / 0 |
| F2 | fill | Missing value stops the fill and writes nothing | PASS | 1 / 1 |
| F3 | fill | Numbers and booleans are converted as text | PASS | 0 / 0 |
| F4 | fill | Allow unfilled placeholders and report them | PASS | 0 / 0 |
| T1 | track-changes | List an insertion and a deletion | PASS | 0 / 0 |
| T2 | track-changes | Accept all tracked changes | PASS | 0 / 0 |
| T3 | track-changes | Reject all tracked changes | PASS | 0 / 0 |
| T4 | track-changes | Refuse to accept a formatting change | PASS | 2 / 2 |
| M1 | comments | Add a comment anchored to exact text | PASS | 0 / 0 |
| M2 | comments | Ambiguous anchor without --occurrence | PASS | 1 / 1 |
| M3 | comments | List comments with anchors and authors | PASS | 0 / 0 |
| V1 | validate | Validate a good file | PASS | 0 / 0 |
| V2 | corrupt-file | Validate a truncated file | PASS | 1 / 1 |
| S1 | summary | Structure summary of a small report | PASS | 0 / 0 |
| PDF1 | to-pdf | Convert to PDF with Word or LibreOffice, or refuse when neither is installed | PASS | 2 / 2 |
| X1 | out-of-scope | Legacy .doc file is refused | PASS | 2 / 2 |
| X2 | out-of-scope | Password-protected file is refused | PASS | 2 / 2 |
| X3 | out-of-scope | Macro-enabled .docm is refused | PASS | 2 / 2 |
| X4 | missing-file | Missing input file | PASS | 1 / 1 |
| X5 | safety | Existing output is not replaced without --force | PASS | 1 / 1 |
| X6 | safety | Output equal to input is refused without --in-place | PASS | 1 / 1 |
| X7 | out-of-scope | Request for a command the tool does not have | PASS | 1 / 1 |
