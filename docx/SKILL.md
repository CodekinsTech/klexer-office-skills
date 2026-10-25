# DOCX skill

Read, create and change Word (.docx) files from the command line. One script does everything:

    python scripts\docx_tool.py <command> <input> [options]

Run it from any folder. Paths can be absolute or relative to the folder you run from.
The script never asks questions and never reads from stdin.

## Rules for every command

- Exit 0 means done and the output was re-opened and checked. Exit 1 means an error (bad input, bad option, a file that already exists, or a failed check); nothing is written. Exit 2 means the file or request is out of scope; nothing is written.
- Errors go to stderr as one line of text. Normal output goes to stdout.
- `--json` (where listed) prints one JSON object to stdout instead of text.
- Commands that write a file take `--out PATH`. Without `--out`, the output is `<input name>.<command>.docx` in the same folder as the input. For example `report.docx` with `replace` gives `report.replace.docx`. The output is never the input unless you pass `--in-place`.
- If the output file already exists, the command exits 1. Add `--force` to replace it.
- `--in-place` changes the input file itself. It is never the default. It cannot be combined with `--out`.
- The output folder must already exist.
- Out of scope, exit 2: legacy `.doc`, password-protected files, macro-enabled `.docm` (the macros are never removed).

Out-of-scope or error messages you may see (exact text):
- `Input file not found: <path>` (exit 1)
- `File is not a readable .docx (corrupt or truncated): <path>` (exit 1)
- `Output already exists: <path>. Choose another --out or pass --force.` (exit 1)
- `Refusing to overwrite the input. Use --in-place to change it.` (exit 1)
- `Legacy .doc is not supported. Save the file as .docx and try again.` (exit 2)
- `This Word file is password-protected. The skill does not open protected files.` (exit 2)
- `Macro-enabled .docm files are not supported. Save as .docx (macros are dropped).` (exit 2)

## Commands

### read
Prints the body paragraphs and tables in document order.
- Input: a .docx. Options: `--json`, `--no-tables`.
- Text output: one line per paragraph. Each table row is one line with its cells joined by a tab, and a blank line follows each table.
- JSON output: `{"paragraphs": [{"style": "...", "text": "..."}], "tables": [[["cell", ...], ...]], "order": [{"kind": "paragraph", "index": 0}, {"kind": "table", "index": 0}], "headers": ["..."], "footers": ["..."]}`. Headers and footers are returned as plain text, one string per header or footer part.
- Exit: 0 read; 1 missing or corrupt; 2 legacy, encrypted or macro-enabled.
- Example: `python scripts\docx_tool.py read contract.docx --json`

### create
Makes a new .docx from a plain text outline.
- Input: a UTF-8 `.txt` outline. Options: `--out`, `--force`, `--json`.
- Outline rules: `# text` is Heading 1; `## text` is Heading 2; `- text` is a List Bullet; `1. text` is a List Number; any other non-blank line is a Normal paragraph; blank lines separate blocks and make no paragraph.
- Output: a new .docx with the standard Word styles.
- JSON output: `{"paragraphs": N, "output": "<path>"}`.
- Exit: 0 created and re-opened; 1 outline missing, empty (no text lines), or not UTF-8; 2 none.
- Example: `python scripts\docx_tool.py create outline.txt --out letter.docx`

### replace
Finds exact text and replaces it.
- Input: a .docx. Options: either `--find TEXT --replace TEXT`, or `--pairs FILE.json` (a JSON object that maps each find text to its replacement). Plus `--out`, `--force`, `--json`, `--in-place`.
- Matching is case-sensitive. A match may span several runs in one paragraph; the new text goes into the first run the match touches, so that run's formatting is kept.
- Searched: body paragraphs, table cells (nested tables too), text boxes, and all headers and footers. Not searched: footnotes, endnotes, comments and field codes (these are listed in the report).
- Tracked changes: text inside a tracked deletion is ignored. A match that touches a tracked insertion is not changed and is counted in `skipped_in_tracked_change`.
- With `--pairs`, the pairs run in the order they appear in the JSON, and each one sees the result of the ones before it.
- Zero matches is not an error: the copy is still written and the count is 0.
- JSON output: `{"replacements": N, "per_pair": {"find": n}, "skipped_in_tracked_change": N, "remaining_after": {"find": n}, "output": "<path>", "not_searched": ["footnotes", "endnotes", "comments", "field codes"]}`.
- Exit: 0 done and verified; 1 missing options, empty find text, bad JSON, or the usual file and output errors; 2 out of scope input.
- Example: `python scripts\docx_tool.py replace report.docx --find ACME --replace Globex --json`

### fill
Fills `{{name}}` placeholders from a JSON file. Placeholder names use letters, digits and underscores only, and must have no spaces inside the braces.
- Input: a .docx template and `--values FILE.json` (a JSON object of name to value). Plus `--out`, `--force`, `--json`, `--allow-unfilled`, `--in-place`.
- Values: JSON strings, numbers and booleans are converted with Python `str()` (so `true` becomes `True`). Lists and objects are refused. A value may not contain `{{`.
- Same search scope and tracked-change rules as `replace`.
- If any placeholder has no value, nothing is written and the message lists the names, for example `Missing values for: DATE, TOTAL. No file written.` With `--allow-unfilled`, the file is written and the unfilled placeholders stay as they are.
- JSON output: `{"filled": N, "per_name": {"NAME": n}, "unfilled": ["DATE"], "skipped_in_tracked_change": N, "output": "<path>", "not_searched": [...]}`.
- Exit: 0 filled and verified; 1 unfilled placeholders (without `--allow-unfilled`), bad JSON, or a list, object or null value; 2 out of scope input.
- Example: `python scripts\docx_tool.py fill letter.docx --values data.json --out letter_final.docx`

### track-show
Lists tracked changes.
- Input: a .docx. Option: `--json`.
- One row per revision: `id`, `kind` (`insert`, `delete` or `unsupported`), `author`, `date`, `text`, `paragraph_index`, `location` (`body`, `header` or `footer`), `detail` (null, or the reason for `unsupported`; `paragraph mark` for a new or deleted paragraph mark, which is resolved like other insertions and deletions, with empty `text`). Formatting changes, moves, and table row or cell changes are listed as `unsupported` so they are visible but not changed.
- JSON output: `{"revisions": [ ...rows... ]}`. Zero revisions gives an empty list.
- Exit: 0 listed (zero is fine); 1 missing or corrupt; 2 legacy, encrypted or macro-enabled.
- Example: `python scripts\docx_tool.py track-show redline.docx --json`

### track-resolve
Accepts or rejects tracked insertions and deletions.
- Input: a .docx. Options: `--mode accept|reject` (required), and either `--id N` (repeatable) or `--all`. Plus `--out`, `--force`, `--json`, `--in-place`.
- Rules: accept an insertion keeps its text. Accept a deletion removes its text. Reject an insertion removes its text. Reject a deletion keeps its text as normal text.
- Paragraph marks: a removed mark joins its paragraph with the next one, as Word does. A section break (a removed mark whose paragraph holds the section properties) is removed the same way when the section before and the section after it have the same page layout; the section after keeps its own properties. Refused (exit 2, nothing written): a break whose layout differs from the section after it, a tracked change to section properties, and a removed mark with no following paragraph in its story or cell.
- If the chosen set includes any `unsupported` revision, nothing is written and the command exits 2. Use `track-show` to see them. The one exception is a paragraph-property change that Word pairs with a deleted paragraph mark: it is resolved with that mark when both are chosen (`--all`), and choosing the mark alone exits 2.
- The output is checked again: the number of tracked changes left must equal the number not chosen.
- JSON output: `{"resolved": N, "remaining": N, "output": "<path>"}`.
- Exit: 0 resolved and verified; 1 bad id (`No tracked change with id N.`) or missing mode or selection; 2 an unsupported revision is in the chosen set.
- A comment whose text is all removed by the resolution is removed too, as Word does (for example a comment on text that is deleted and then accepted). Comments that keep some text are kept.
- Example: `python scripts\docx_tool.py track-resolve redline.docx --mode accept --all --out clean.docx`

### comments-list
Lists the comments.
- Input: a .docx. Option: `--json`.
- One row per comment: `id`, `author`, `initials`, `date`, `anchor_text` (the text between the comment's start and end markers), `text`.
- JSON output: `{"comments": [ ...rows... ]}`.
- Exit: 0 listed (zero is fine); 1 missing or corrupt; 2 legacy, encrypted or macro-enabled.
- Example: `python scripts\docx_tool.py comments-list review.docx --json`

### comments-add
Adds a comment anchored to exact text in the document body.
- Input: a .docx. Options: `--anchor TEXT` (required), `--text TEXT` (required), `--author NAME` (required, no default), `--initials XX` (optional), `--occurrence N`, plus `--out`, `--force`, `--json`, `--in-place`.
- The anchor is searched in body paragraphs, including table cells and text boxes, but not in headers or footers. A match may span runs. The runs are split at the anchor edges so formatting is kept.
- If the anchor occurs more than once and `--occurrence` is not given, nothing is written and the message gives the count, for example `Anchor text occurs 2 times. Pass --occurrence 1..2.`
- The output is re-opened and must list the new comment with the same anchor, text and author.
- JSON output: `{"comment_id": N, "anchor_text": "...", "text": "...", "author": "...", "output": "<path>"}`.
- Exit: 0 added and verified; 1 anchor not found, ambiguous, or the usual file errors; 2 out of scope input.
- Example: `python scripts\docx_tool.py comments-add contract.docx --anchor "signed" --text "Check date" --author "Review Bot" --out reviewed.docx`

### validate
Checks that a .docx opens and that its XML is intact.
- Input: a .docx. Option: `--json`.
- Checks, in order: the file is a readable zip; `[Content_Types].xml` and `word/document.xml` exist; every `.xml` and `.rels` part parses; every internal relationship target exists in the file; python-docx opens the file; no `w:delText` outside a `w:del`; no tracked change inside another tracked change. Each check is `PASS`, `FAIL` or `SKIP`. The checks after a failed zip check are skipped.
- Text output: one line per check, then `valid` or `invalid`. JSON output: `{"checks": [{"name": "...", "passed": true, "detail": "..."}], "verdict": "valid"}`.
- Exit: 0 valid; 1 invalid or file missing; 2 legacy, encrypted or macro-enabled (these cannot be checked).
- Example: `python scripts\docx_tool.py validate report.docx`

### to-pdf
Converts a .docx to PDF with Microsoft Word when Word is installed (Windows, with pywin32). Otherwise it uses LibreOffice.
- Input: a .docx. Options: `--out PATH` (must end in `.pdf`), `--force`, `--timeout SECONDS` (default 120).
- With Word: Word opens a copy of the file in a temporary folder inside the output folder and saves it as PDF in a separate process, so `--timeout` can stop it. The PDF shows the document as Word displays it (tracked changes shown as markup). The input is not changed.
- Without Word: finds `soffice` on PATH, or in the default LibreOffice install folder, and runs it headless in a temporary folder inside the output folder.
- The PDF is re-opened and must have at least one page. The output line says which converter was used.
- Default output: `<input name>.to-pdf.pdf`.
- Exit: 0 converted and verified; 1 input or output problem, timeout (`PDF conversion timed out after N seconds.`), or no PDF was produced; 2 neither Word nor LibreOffice is installed (`LibreOffice is not installed. Install it or convert the file in Word.`), or the input is legacy, encrypted or macro-enabled.
- Example: `python scripts\docx_tool.py to-pdf letter.docx --out letter.pdf --timeout 300`

### summary
Describes the structure of a document.
- Input: a .docx. Option: `--json`.
- JSON output: `{"headings": [{"level": 1, "text": "..."}], "paragraph_count": N, "table_count": N, "table_sizes": [[rows, cols]], "styles_used": {"Normal": N}, "section_count": N, "has_comments": false, "has_tracked_changes": false}`. Headings are paragraphs styled Heading 1 to Heading 9. `paragraph_count` counts body paragraphs, headings included, and table cell paragraphs are left out. `styles_used` counts the same paragraphs, so its counts add up to `paragraph_count`.
- Text output: the same facts as a short outline.
- Exit: 0 summarized; 1 missing or corrupt; 2 legacy, encrypted or macro-enabled.
- Example: `python scripts\docx_tool.py summary report.docx --json`

## Requirements

Python 3.11 on Windows, macOS or Linux, with the packages in `scripts\requirements.txt`. `to-pdf` needs Microsoft Word (Windows, with pywin32 installed) or LibreOffice installed. See README.md for details, decisions and the list of what is not done.
