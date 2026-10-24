# DOCX v2 accuracy benchmark

Runs 30 tasks through the v2 docx CLI (`docx\scripts\docx_tool.py`) and scores each one against a hand-written expected result.

## How to run

    "<venv>\Scripts\python.exe" run_benchmark.py
    "<venv>\Scripts\python.exe" run_benchmark.py --only P1 T4

The venv needs python-docx, lxml and pypdf. The script rebuilds the inputs in `work\` on every run (safe to delete), runs each CLI call with that task's folder as the working directory, and writes `results.md` next to the script. Options: `--python`, `--work`, `--out`, `--timeout` (default 180 s per call), `--only`.

## Files

- `make_tasks.py`: the 30 task definitions, the input builders, and the readers that check output files.
- `run_benchmark.py`: runs the CLI, compares results with the expectations, writes `results.md`.
- `results.md`: generated report with the score, per-task PASS/FAIL, diffs for failures, and the score by category.

## Categories

| Category | Tasks | What it covers |
|---|---|---|
| read | R1 | Paragraphs and tables in document order (text mode) |
| create | C1-C3 | Outline syntax, UTF-8, empty outline refused |
| replace | P1-P4 | Body, table and header text; matches across runs; `--pairs` order and zero matches; tracked insertions and deletions |
| fill | F1-F4 | Placeholders in body, table and footer; missing values; value types; `--allow-unfilled` |
| track-changes | T1-T4 | `track-show`; accept all; reject all; formatting change refused |
| comments | M1-M3 | `comments-add` with exact anchor; ambiguous anchor refused; `comments-list` |
| validate | V1 | Good file passes every check |
| corrupt-file | V2 | Truncated file is reported invalid |
| summary | S1 | Headings, counts, table sizes, styles |
| to-pdf | PDF1 | Conversion with Word (or LibreOffice), or the refusal when neither is installed |
| out-of-scope | X1-X3, X7 | Legacy .doc, password-protected file, macro-enabled .docm, unknown command |
| missing-file | X4 | Input path does not exist |
| safety | X5-X6 | Existing output is not replaced without `--force`; output equal to input is refused |

## Expected-result method

1. Every expectation was written from `specs\docx.md` and `docx\SKILL.md` before the benchmark was run. Nothing was taken from a run of the tool.
2. Expectations are literal values: exit codes, JSON values and counts, exact stdout text for `read`, error-message substrings from the spec's error table, and the exact body text and styles of output files.
3. Output files are checked by re-reading them with python-docx, zipfile and lxml, not by trusting the tool's own report.
4. Every task also checks that no file was written unless the task lists it under `new_files`, and that no file which existed before the run changed unless the task lists it under `changed`. This catches overwrites of the input and partial writes on refusals.
5. Where the result depends on the machine (Word or LibreOffice installed or not), the task holds one expectation per converter: `word` (Word is registered and the benchmark's Python imports pywin32), `libreoffice` (PATH first, then `C:\Program Files\LibreOffice\program\soffice.exe`), or `none`. The results header says which one was used.
6. Where the spec and SKILL.md give different JSON shapes, SKILL.md was followed, because it is the CLI reference. For example, `track-show` returns `{"revisions": [...]}`, and `replace` also reports `remaining_after`.

## Scoring

Score = passed / runnable. A task is PASS when every check holds, and FAIL when any check does. A task is NOT RUNNABLE only when its fixture cannot be built, which is a benchmark bug and is listed separately. Tool failures are never NOT RUNNABLE. Scores by category are in `results.md`.

## Fixtures and their limits

- Input documents are built with python-docx. Tracked insertions and deletions, property changes (`w:pPrChange`) and comment anchors are added as raw OOXML, because python-docx has no API for them.
- The legacy `.doc` and encrypted `.docx` inputs are structural OLE compound files, written by `_cfb` in `make_tasks.py` from the MS-CFB layout. The `.doc` has a `WordDocument` stream. The encrypted file has `EncryptionInfo` and `EncryptedPackage` streams, the streams Word writes for protected OOXML. Neither file was saved by Word, and no encryption library is installed.
- Fixture history, for transparency: the first version of the encrypted fixture was a bare 512-byte OLE header. The CLI reported "Legacy .doc" for it, so X2 failed. The fixture was changed to contain the `EncryptionInfo` and `EncryptedPackage` streams, and X2 then passed. A failure from the bare-header version was a fixture problem, not evidence about the tool.
- Tracked changes were written by hand from the OOXML element model, not produced by Word. The spec's release notes call for Word-generated revisions before release; that set is still missing.

## Coverage gaps (not among the 30)

- `--in-place` (never changes the input without it, and changes it with it)
- `track-resolve --id N` and the "No tracked change with id N" error
- `comments-add --occurrence N` success path
- Missing required options (`--replace`, `--author`, `--mode`) naming the option
- `read --no-tables` and headers and footers in `read --json`
- `summary` with `has_comments` or `has_tracked_changes` set to true
- `validate` on a w:delText outside a w:del, and on nested tracked changes
- `to-pdf` through LibreOffice: LibreOffice is not installed where this was built, so only the Word branch of PDF1 ran
- Footnotes and endnotes: only the `not_searched` list is checked, not a footnote replacement
