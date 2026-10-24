# PPTX v2 accuracy benchmark

Measures the v2 `pptx` skill (`scripts/pptx_tool.py`) against 30 tasks a business
user might ask a bot to do with a deck. Each task says what the user asked, the
exact CLI command, and the exact result the spec requires. The runner checks the
tool's exit code, messages, JSON, output decks, folder contents and input files.

## Files

- `make_tasks.py` builds `cases/<task_id>/`. Each folder has the input files
  (decks made with python-pptx, outlines, JSON data, expect files, PNGs) and
  `task.json`, the task definition.
- `run_benchmark.py` runs every task through the CLI in a scratch copy under
  `_work/`, compares the result with `task.json`, and writes `results.md`.
- `results.md` is the report: score, score by category, top failure categories,
  per-task PASS/FAIL, and for each failure a diff of expected against actual plus
  the stdout and stderr.

## How to run

Use the venv that has python-pptx (`requirements.txt` of the skill):

    VENV="<your-venv-folder>\venv-v2-pptx\Scripts\python.exe"
    cd "<skill-folder>\pptx\benchmark"
    "$VENV" make_tasks.py
    "$VENV" run_benchmark.py

Options for `run_benchmark.py`:

- `--only TEXT` runs only the task ids that contain TEXT, for example `--only P1`.
- `--tool PATH` points at another copy of `pptx_tool.py`.
- `--python PATH` is the interpreter that runs the CLI.

The runner reads output decks with python-pptx directly, not through the tool.
Run it with the same venv, because it imports python-pptx itself.

## Categories

| Category | Tasks | What it covers |
|---|---|---|
| create | 3 | Outline with levels and notes; text before the first title is refused; Thai, Arabic and accented text round trip |
| read | 3 | Column chart JSON; table and `--slides` subset; reading never creates a notes page or changes the file |
| replace | 3 | Title, table cell, split-run text and notes, with per-slide counts; `--pairs` file; default output name and input never overwritten |
| add | 2 | Insert at a position with levels, plain text and notes; position past the end is refused |
| delete | 1 | Delete a slide; its text is gone from every part of the package |
| reorder | 1 | Permutation `3,1,2` |
| duplicate | 2 | Copies text, notes and layout; a slide with a chart is refused and the message names the chart |
| split | 2 | `--every 2` gives three parts with the right titles; ranges that leave a slide out are refused |
| replace-image | 1 | Swap the picture; new image hash, same box, same name |
| chart-data | 3 | Update data with type and series colour kept; value count mismatch refused; XY chart refused |
| check-text | 2 | Found and missing phrases with JSON; whitespace normalisation; chart categories and series names |
| validate | 2 | A good deck passes all 9 checks; a deck cut to 1000 bytes fails |
| refusal | 5 | Corrupt file, missing file, legacy `.ppt`, macro-enabled `.pptm`, out-of-scope `render` request |

The score is passed/runnable, where runnable means the CLI started and the task
was checked. A task that cannot run is listed under "Not runnable" and left out
of the score.

## How expected results are made

Every expected value is written by hand in `make_tasks.py` from `specs/pptx.md`
and `SKILL.md`. Nothing is recorded from a run of the tool. The inputs are built
so the answer can be worked out in advance, for example "ACME" in a title, a
table cell, a two-run split and the notes gives 4 replacements, 1 in the notes,
and per-slide counts 1: 2 and 2: 2.

Where the spec leaves a detail open, the benchmark makes these choices:

- Notes are compared after whitespace is collapsed, because the spec joins several
  `Notes:` lines with new lines without saying how they are stored.
- Body paragraphs are reported with a 1-based level (level 1 is the top bullet).
- A chart type is matched in lower case as a substring (`column`), since the spec
  does not name the type strings.
- Errors are checked by the exact message text from the spec's error table where
  the spec gives one. Where it gives none, the check only requires a non-empty
  stderr or the exit code.
- `read --slides` checks that `slide_count` counts the whole deck.
- Checks on JSON keys use these forms in `task.json`: `"a.b.0.c"` exact, `"len:key"`
  length, `"~key"` case-insensitive substring, `"all:list.field"` every item equal.

Every task also checks that the input files are byte-identical after the run,
and that a failed command wrote no output file.

## Known gaps

- Not covered: SmartArt, video or audio duplication, encrypted decks, a `--template`
  case, a multi-slide `replace-image` shape index, and animations or themes
  (out of scope in the spec).
- `replace-image` and `reorder` have one task each, and `delete` has one.
  This follows the 25 to 30 task limit.
- `X5_refusal_out_of_scope_render` expects exit 2 and the `Not supported: rendering
  slides to PDF or images.` message. `render` is a feature the tool refuses (exit 2);
  an unknown command name is a usage error (exit 1).
