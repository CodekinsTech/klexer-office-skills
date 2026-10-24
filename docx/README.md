# DOCX skill - developer notes

This folder holds the v2 docx skill: one command-line tool (`scripts\docx_tool.py`), its tests, and the instructions for AI agents (`SKILL.md`). The behaviour comes from `specs\docx.md` in the parent folder. This file records libraries, how to test, the decisions made where the spec left a detail open, and what is not done.

## Layout

- `SKILL.md` - command reference for agents (inputs, outputs, exit codes, examples).
- `scripts\docx_tool.py` - the whole tool. One entry point, subcommands as the spec lists them.
- `scripts\requirements.txt` - pinned packages.
- `tests\test_docx_tool.py` - one pytest test per acceptance test in spec section 3. Each test builds its input files with python-docx.
- `tests\test_word_com.py` - checks against files Word made (COM automation). Skipped where Word is not installed.

## Libraries and licences

Checked against the licence metadata of the installed packages.

| Package | Version used | Licence | Used for |
|---|---|---|---|
| python-docx | 1.2.0 | MIT | opening, reading, saving, comments API (`Document.add_comment`) |
| lxml | 6.1.3 | BSD-3-Clause | XML edits (revisions, run splitting, validation) |
| pypdf | 6.21.2 | BSD-3-Clause | checking the PDF made by `to-pdf` (imported only by `to-pdf`) |
| pytest | 9.1.1 | MIT | tests only |
| pywin32 | 312 | PSF licence | `to-pdf` through Word, and `tests\test_word_com.py`; imported only there |
| Microsoft Word | 16.0 (build 20430) | Microsoft Office licence (external program) | `to-pdf` (preferred converter) and the Word checks; driven through COM, never bundled |
| LibreOffice | not installed here | MPL-2.0 (external program) | `to-pdf` fallback when Word is missing; run as a separate program, never imported |

The LibreOffice licence still needs checking against the LICENSE file of the version a user installs, as the spec asks. That has not been done.

## Running the tests

Environment used: Windows 11, Python 3.11.9, a venv at `<your-venv-folder>\venv-v2-docx`.

    python -m venv venv
    venv\Scripts\python.exe -m pip install -r scripts\requirements.txt
    venv\Scripts\python.exe -m pytest tests -v

The `to-pdf` test checks the PDF branch when Word or LibreOffice is available, and the exit 2 branch when neither is. On this machine Word is the converter, so the Word branch ran.

The Word checks need the same venv with pywin32 (`venv\Scripts\python.exe -m pip install pywin32==312`):

    venv\Scripts\python.exe -m pytest tests\test_word_com.py -v

They start Word through COM, build ten Word-made files (six general, four paragraph-mark cases), run the write commands on them, and open each output in Word. A prompt from Word while it opens a file, or an exception from Word, fails the check. They take about three minutes. Word's `UseLocalUserInfo` option is set for the run so that revision and comment authors are the names the test uses; the old value is restored before Word quits.

## Decisions where the spec was open

Output and JSON shape
- Every write command adds an `output` key to its JSON. `replace` and `fill` also add `not_searched`, as the spec says they must list it. `replace` adds `remaining_after` (how many matches of each find text are still in the output).
- Row-based JSON is wrapped in one object: `{"revisions": [...]}`, `{"comments": [...]}`, `{"checks": [...], "verdict": ...}`.
- `read` JSON `order` is a list of objects `{"kind": "paragraph"|"table", "index": n}`.
- `read` JSON `headers` and `footers` are lists of plain-text strings, one per header or footer part, sorted by part name. Text mode for `read` shows the body only.
- `track-show` rows add `location` (`body`, `header` or `footer`) and `detail` (why a row is `unsupported`).

Output names and folders
- The default output name is `<input stem>.<command>.docx`. For `create` the input is the outline, so `outline.txt` gives `outline.create.docx`.
- `to-pdf` defaults to `<input stem>.to-pdf.pdf`.
- The output folder must exist. The tool does not create it.
- `--in-place` with `--out` is a usage error (exit 1).

Reading and matching
- The text the tool sees is the "accepted view": tracked insertions are included, tracked deletions are ignored. `read`, `summary`, `replace` and `fill` all use it.
- A match or placeholder that touches a tracked insertion is not changed, and it is counted as `skipped_in_tracked_change`. A match never runs into deleted text.
- Matches are found left to right without overlap. Replacements are applied right to left inside a paragraph, so earlier positions stay valid.
- `--pairs` runs the pairs in JSON order. Later pairs see the results of earlier ones, so chained replacements are possible.
- A replacement empty string is allowed (it deletes the match). Only the find text may not be empty.
- Header and footer parts are found from the package, so all of them are searched, including any not referenced by a section.
- Table cell text in `read` joins the cell's paragraphs with a newline. Text mode turns those newlines into spaces so each row stays on one line.

Tracked changes
- Plain revisions are `w:ins` and `w:del` whose parent is paragraph content (`w:p`, `w:hyperlink`, `w:smartTag`, `w:sdtContent`, `w:customXml`, `w:fldSimple`, `w:dir`, `w:bdo`). Paragraph-mark revisions (inside `w:pPr/w:rPr`) are plain too; see Paragraph marks.
- These are `unsupported` (never changed, exit 2 when chosen): formatting changes (`w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, `w:tblGridChange`, `w:numberingChange`), except the paragraph-property change that Word pairs with a paragraph-mark revision (see Paragraph marks below); changes to table rows and cells; moves (`w:moveFrom`, `w:moveTo`, text in `w:moveFrom` is hidden like a deletion and text in `w:moveTo` counts as an insertion); a revision nested inside another; and a revision with no valid `w:id`.
- Revision ids are the `w:id` values. If two revisions share an id, `--id` selects both.
- `track-resolve` checks that the number of plain revisions left equals the number not chosen.
- A new paragraph typed with tracking on is one revision in Word: the paragraph mark and its text. Here the paragraph mark is its own row, so `track-show` lists one more row than Word's count. The mark row has `kind` `insert` (or `delete`), empty `text` and `detail` `paragraph mark`. Word's count is kept in the tests by leaving those rows out of the comparison.
- Accept and reject match Word's own Accept All and Reject All, including comments (`test_word_com.py`).

Paragraph marks
- A paragraph mark revision is a `w:ins` or `w:del` inside `w:pPr/w:rPr`. Accepting an inserted mark or rejecting a deleted mark keeps the paragraph, and only the marker goes.
- Removing a mark (rejecting an inserted mark, or accepting a deleted one) joins the paragraph with the next paragraph in the same story or table cell. The joined content moves to the start of the next paragraph. Word's rules, checked against Word's own Accept All and Reject All on Word-made files (`test_word_com.py`, `PM_FIXTURES`):
  - A paragraph with no content left (for example a whole paragraph deleted with Track Changes) is removed. The next paragraph keeps its own properties.
  - Accepting a deleted mark of a paragraph that still has text: the joined paragraph takes the first paragraph's properties (`Title` in Heading 1 plus `Body` in Normal gives `TitleBody` in Heading 1).
  - Rejecting an inserted mark: the joined paragraph keeps the next paragraph's properties.
- Word writes a paragraph-property change (`w:pPrChange`) with a mark deletion when the joined paragraph's formatting changes. That change is resolved together with the mark: accept drops it, reject restores the old properties. It is resolved only when it sits on a paragraph whose mark is in the same resolve, and it must be selected too (with `--all`, or with `--id` for both). Selecting only the mark is refused with exit 2 and the message "must be selected too".
- Refused with exit 2 and nothing written (`cannot be resolved by this skill`):
  - a mark that ends a section (a section break);
  - a mark that must be removed with no paragraph after it in the same story or table cell (for example the last paragraph of the body);
  - a formatting change on a paragraph involved in a mark removal that is not resolved with it;
  - any other `unsupported` revision, as before (`rPrChange`, table row or cell changes, moves, nested revisions). Comment markers inside a removed element are moved out first. A comment whose text is all removed is dropped with its entry in `comments.xml`, because Word does the same.

Summary
- Headings come from paragraph styles named `Heading 1` to `Heading 9` (not `Title`).
- `paragraph_count` and `styles_used` count body-level paragraphs, including paragraphs inside block-level content controls. Table cell paragraphs are left out.
- `table_count` counts top-level tables. `table_sizes` gives rows and the number of `w:gridCol` entries.
- `has_comments` is true when the comments part has at least one comment. `has_tracked_changes` is true when `track-show` would list any row.

Fill
- Values: `str()` of the JSON value, so `true` becomes `True`. Values containing `{{` are refused with exit 1, because they could create new placeholders. The message is `Value for NAME must not contain "{{".`
- A placeholder that is in a tracked insertion, or that has no value, is listed under `unfilled`.
- Newlines inside a value are written as plain text characters, not as line breaks. Word displays them as spaces. Not changed, because the spec does not ask for it.

Comments
- The anchor is searched in body paragraphs only (including table cells and text boxes), not in headers or footers.
- The occurrence count covers the whole body. `--occurrence` picks one.
- The comment is added with python-docx `Document.add_comment`, which sets the date to the current UTC time. If `--initials` is not given, the `initials` attribute is left out.
- The anchor is matched across runs. Runs are split at the anchor edges (new runs copy the original run properties).
- Comments in headers and footers are not supported.
- `comments-add` moves the new comment's markers out of any tracked insertion or deletion around them, and puts the reference run right after the end marker (the layout Word writes). Without this, rejecting the insertion made Word show no comment at all.

Create (outline)
- A UTF-8 file with a BOM is accepted. Each line is stripped before it is read.
- Markers: `## ` (Heading 2) is checked before `# ` (Heading 1). `- ` is List Bullet. A line matching `1. text` (any number) is List Number. Anything else, including `### text` and `1.text`, is Normal text with the marker kept.
- The output uses the default template from python-docx. Its core document properties are left as python-docx sets them. The default template has no tables.

Validate
- Validate runs its own checks, so a corrupt file gives a readable `FAIL` list (exit 1) instead of the generic message. Legacy, encrypted and macro-enabled files still exit 2 first.
- XML is parsed with entity resolution and network access turned off.
- The package check uses `zipfile.testzip()`, so a bad CRC is also reported.

Macro and legacy detection
- `.doc` and `.docm` names are refused by extension. A file whose first bytes are the OLE header is legacy, or encrypted when it contains the `EncryptionInfo` stream name. A zip whose `[Content_Types].xml` says macro-enabled is refused as `.docm` even if it has a `.docx` name.

To-pdf
- Word is used when Word is registered and pywin32 imports. Word opens a copy of the input in the temporary folder, so the input is not locked or changed. Word saves the copy as PDF (format 17) in a child process, hidden as `_word-pdf` (not listed in `--help`), so `--timeout` can stop it.
- After a timeout, the Word instances that started during the conversion are stopped; an instance that was already running is not touched. The stop is best effort: an instance Word starts after its child is already killed can survive.
- Without Word, LibreOffice is used. It gets a private profile folder inside the output folder and writes its log to a file, so a running LibreOffice does not block the call. A child `soffice.bin` may keep running after a timeout on Windows (not solved).
- With neither program, exit 2 with the spec's message.
- The PDF shows the document as Word displays it: tracked changes appear as markup, deleted text included (checked by reading the PDF text).
- Word cannot re-open the PDF for checking here: opening a PDF in Word waited on a prompt that COM never returned from. The tests therefore compare the page count with Word's and look for the document's first words in the PDF text.
- `--out` must end in `.pdf` (exit 1 otherwise). pypdf is imported only when `to-pdf` runs, so the other commands work without it.

Locks and writes
- Each write goes to a temp file in the output folder, is re-opened and checked, and only then is moved into place with `os.replace`. A failed check deletes the temp file.
- `PermissionError` at any point gives the locked-file message (exit 1).

## Verification after writing

| Command | What the re-open check tests |
|---|---|
| create | paragraphs and their styles equal the outline |
| replace | paragraph count is unchanged; `remaining_after` is reported, not enforced |
| fill | placeholders left in the output equal the ones left behind |
| track-resolve | plain tracked changes left equal the number not chosen |
| comments-add | the new comment is listed with the same anchor, text and author |
| to-pdf | the PDF opens and has at least one page |

## Not done

- LibreOffice conversion has not been run here (LibreOffice is not installed). Its branch runs only when Word is missing.
- `read`, `summary`, `replace` and `fill` do not join a paragraph whose mark is deleted. Word's final view shows such a paragraph joined with the next one, so `read` lists two paragraphs where Word shows one. This is why the four paragraph-mark fixtures are kept out of the read and summary checks. `track-resolve` handles them.
- Word-made files: ten files, built by Word through COM in `tests\test_word_com.py`, not by typing in the Word window. Only Word 16.0 was used. Word may write shapes these files do not cover (footnotes, text boxes, tracked moves). The hand-built XML tests in `test_docx_tool.py` remain.
- Footnotes, endnotes, comments (as text to search) and field codes are not searched. The report lists them.
- Images, mail merge, and layout are not supported (out of scope in the spec).
- Content controls are read and searched only as plain text. Their own properties are not handled.
- Only Windows with Python 3.11.9 has been run. macOS and Linux code paths (the LibreOffice location only) are written but not run.
- LibreOffice licence check (see Libraries and licences).
- Value line breaks in `fill` (see Decisions).

## Acceptance test results

Run on Windows 11, Python 3.11.9, Word 16.0, python-docx 1.2.0, lxml 6.1.3, pypdf 6.21.2, pywin32 312, pytest 9.1.1: 85 tests pass (25 in `test_docx_tool.py`, 60 in `test_word_com.py`). The benchmark (`benchmark
un_benchmark.py`, run with the venv that has pywin32) scores 30/30; its to-pdf task ran the Word variant.
