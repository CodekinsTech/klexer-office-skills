# PDF skill

One command-line tool, `scripts/pdf_tool.py`, for reading, checking, filling and editing PDF files.
Run it from any folder. Give absolute paths or paths relative to the current folder.

    python "<skill folder>\scripts\pdf_tool.py" COMMAND [options]

Install the libraries once: `pip install -r scripts/requirements.txt`.

## Rules that apply to every command

- **Exit codes:** `0` done (and checked again, for commands that write a file). `1` bad input, bad option, or a failed check. `2` a password is needed, or the job is out of scope (no text layer, no form fields, a character the font cannot draw).
- **Output:** a read command prints one JSON object with `--json`, or readable text without it. Errors always go to stderr as a short message. Nothing is ever prompted for.
- **`--out FILE`** on a read command writes the same result to a new file. It refuses an existing file unless `--force` is given. Nothing is printed to stdout in that case.
- **Write commands** need `--out FILE` (or `--out-dir` for split). They write a temporary file, re-open it, check it, and only then move it into place. A failed check keeps nothing and exits 1. `--json` prints the result, including `"verified": true`.
- **Inputs are never changed by default.** Writing to an input path with `--out` is refused: `Refusing to overwrite an input. Write to a new file.` The only exception is `--in-place` (fields-fill, rotate, reorder, watermark): the result replaces the input, after it has been checked. No backup is kept. `--in-place` is refused for a password-protected input (`--in-place is refused for a password-protected PDF. Write to a new file with --out.`), so encryption is never removed by accident.
- **File locked:** if another program has the file open with a lock, the message is `The file is open in another program. Close it and try again.` (exit 1). Nothing is changed.
- **Passwords** are read only from an environment variable whose name you pass with `--password-env NAME`. The password is never printed or written to output. If you do not pass `--password-env`, the message names `PDF_PASSWORD`.
- **Page lists** are 1-based and look like `1-3,5,8-` (a trailing dash means "to the last page").

## Commands

### text
Extract text, page by page.
- Inputs: `INPUT.pdf`. Options: `--pages LIST` (default all), `--password-env NAME`, `--json`, `--out FILE`, `--force`.
- Output (JSON): `{"page_count": N, "pages": [{"page": 1, "text": "..."}], "no_text_pages": [...]}`. Pages with no text layer have `"text": ""` and are listed in `no_text_pages`. Text mode prints `--- page N ---` before each page.
- Exit: 0 extracted. 1 file missing, not a PDF, or bad page list. 2 password needed, or no selected page has a text layer (the message says scanned PDFs need OCR, which this skill does not do).
- Example: `python pdf_tool.py text report.pdf --pages 1-2 --json`

### tables
Extract tables from pages (uses pdfplumber with default settings).
- Inputs: `INPUT.pdf`. Options: `--pages LIST`, `--password-env NAME`, `--json`, `--out FILE`, `--force`.
- Output (JSON): `{"tables": [{"page": 1, "index": 1, "rows": [["a", "b"], ["c", "d"]]}]}`. `index` counts tables on that page from 1. Empty cells are `""`. No tables is exit 0 with an empty list.
- Exit: 0 extracted. 1 file or page list error. 2 password needed.
- Example: `python pdf_tool.py tables invoice.pdf --json`

### fields-list
List AcroForm form fields.
- Inputs: `INPUT.pdf`. Options: `--password-env NAME`, `--json`, `--out FILE`, `--force`.
- Output (JSON): `{"fields": [{"name": "FullName", "type": "text", "value": "", "options": [], "page": 1}]}`. `type` is `text`, `checkbox`, `dropdown`, `radio` or `other`. A checkbox `value` is `true` or `false`. `options` lists the allowed values for radio and dropdown fields, and is `[]` otherwise. `page` is the page holding the widget, or `null`. A PDF with no form gives an empty list and exit 0.
- Exit: 0 listed. 1 file error. 2 password needed.
- Example: `python pdf_tool.py fields-list form.pdf --json`

### fields-fill
Fill form fields from a JSON object, then check the saved file.
- Inputs: `INPUT.pdf`. Options: `--values FILE.json` (required; an object of field name to value), `--out FILE` or `--in-place` (one is required), `--password-env NAME`, `--force`, `--json`.
- Value rules: a checkbox takes JSON `true` or `false`. A radio or dropdown takes one of the values listed by `fields-list`. A text field takes a string. Fields not listed are left unchanged. An `other`-type field (list box, push button) cannot be filled.
- Output (JSON): `{"out": "...", "filled": {"FullName": "Ana", "Agree": true}, "verified": true}`.
- Exit: 0 filled and verified. 1 an unknown field name (all names are listed, and no file is written), a value that is not allowed, a bad or missing values file, or a failed check. 2 password needed, or the PDF has no form fields.
- Example: `python pdf_tool.py fields-fill form.pdf --values values.json --out filled.pdf --json`
- `values.json` example: `{"FullName": "Ana", "Agree": true}`

### merge
Join two or more PDFs in the order given.
- Inputs: two or more `INPUT.pdf` paths. Options: `--out FILE` (required), `--password-env NAME` (checked for each input separately), `--force`, `--json`.
- Output: one PDF whose page count is the sum of the inputs. The result says how many pages each input contributed.
- Exit: 0 merged and verified. 1 fewer than two inputs, a missing or corrupt input, or a failed check. 2 an input is password-protected and no password was given.
- Example: `python pdf_tool.py merge a.pdf b.pdf --out all.pdf`

### split
Split a PDF into parts inside a folder.
- Inputs: `INPUT.pdf`. Give exactly one of `--every N` (pages per part) or `--ranges "1-3,4-6"`. Options: `--out-dir DIR` (required; must not exist, or must be empty), `--password-env NAME`, `--json`.
- Output files: `<name>.part01.pdf`, `<name>.part02.pdf`, and so on, inside the folder. With `--ranges`, every page must be covered exactly once.
- Exit: 0 split and verified. 1 both or neither of `--every` and `--ranges` given, a bad range (for example `Ranges must cover each page once. Page 3 is not covered.`), a folder that is not empty, or a file error. 2 password needed.
- Example: `python pdf_tool.py split book.pdf --every 10 --out-dir book_parts`

### rotate
Rotate pages clockwise.
- Inputs: `INPUT.pdf`. Options: `--angle 90|180|270` (required), `--pages LIST` (default all), `--out FILE` or `--in-place` (one is required), `--password-env NAME`, `--force`, `--json`.
- Behaviour: each selected page gets its rotation set to its old rotation plus the angle, modulo 360. Other pages are unchanged.
- Exit: 0 rotated and verified. 1 angle not allowed (`Angle must be 90, 180 or 270.`) or bad page list. 2 password needed.
- Example: `python pdf_tool.py rotate scan.pdf --pages 1 --angle 90 --out scan_r.pdf`

### reorder
Put the pages in a new order. The order must use every page exactly once.
- Inputs: `INPUT.pdf`. Options: `--order LIST` (required, for example `3,1,2`), `--out FILE` or `--in-place` (one is required), `--password-env NAME`, `--force`, `--json`.
- Exit: 0 reordered and verified. 1 the list is not a full permutation (`Order must list every page 1..N exactly once.`). 2 password needed.
- Dropping pages is not allowed here. Use `split` and `merge` for that.
- Example: `python pdf_tool.py reorder deck.pdf --order 3,1,2 --out deck_r.pdf`

### watermark
Draw a stamp or diagonal text over pages.
- Inputs: `INPUT.pdf`. Give exactly one of `--stamp FILE.pdf` (the first page of the stamp file is drawn over each selected page) or `--text "DRAFT"` (diagonal grey text). Options: `--pages LIST` (default all), `--out FILE` or `--in-place` (one is required), `--password-env NAME`, `--force`, `--json`.
- The text watermark uses the Helvetica font, so it can only draw characters that WinAnsi (cp1252) covers.
- Exit: 0 watermarked and verified. 1 neither or both options, or a missing stamp file (`Give exactly one of --stamp or --text.`), or a failed check. 2 password needed, or a character the font cannot draw (`Character U+<hex> on line <N> is not in the font. No text was drawn.`), or the text engine (reportlab) is not installed.
- Example: `python pdf_tool.py watermark memo.pdf --text "DRAFT" --out memo_draft.pdf`

### create
Make a PDF from a UTF-8 plain text file.
- Inputs: `INPUT.txt`. Options: `--out FILE.pdf` (required; must end in `.pdf`), `--font-size N` (default 11), `--page-size A4|Letter` (default A4), `--margin MM` (default 20), `--force`, `--json`.
- Layout: lines wrap to the page width. A blank line adds vertical space. A form feed character starts a new page. A tab is written as four spaces.
- The output is checked: every input line must appear in the output text, in order (spaces and line breaks are ignored in the comparison).
- Exit: 0 created and verified. 1 input missing or not UTF-8, output exists without `--force`, or a failed check. 2 a character the font cannot draw (for example Chinese or Thai text). The message gives the first such character and its line number. Nothing is written.
- Example: `python pdf_tool.py create notes.txt --out notes.pdf --page-size Letter`

### check
Check that a PDF opens and reads cleanly.
- Inputs: `INPUT.pdf`. Options: `--password-env NAME`, `--json`, `--out FILE`, `--force`.
- What it checks: the file opens in strict mode (a broken cross-reference table or trailer is an error), whether it is encrypted, at least one page exists, every page's content decodes, and whether a text layer exists (a warning only).
- Output (JSON): `{"readable": true, "page_count": 3, "encrypted": false, "warnings": [], "errors": []}`.
- Exit: 0 readable with no errors (warnings allowed). 1 corrupt or unreadable (the errors are printed and listed in the JSON). 2 encrypted and no password given (the check is not run).
- Example: `python pdf_tool.py check upload.pdf --json`

### check-text
Check that expected phrases appear in a PDF.
- Inputs: `INPUT.pdf`, `--expect FILE.txt` (required; one phrase per line, blank lines ignored). Options: `--password-env NAME`, `--json`, `--out FILE`, `--force`.
- Matching: whitespace runs are collapsed to one space and matching is case-sensitive. Each phrase is searched across the whole document and the first page where it appears is reported.
- Output (JSON): `{"found": [{"phrase": "Alpha", "page": 1}], "missing": ["Gamma"]}`.
- Exit: 0 all phrases found. 1 one or more missing (the list is printed to stderr as `N phrase(s) missing: ...`). 2 password needed, or no text layer.
- Example: `python pdf_tool.py check-text export.pdf --expect must_have.txt --json`

## Things this skill does not do

- OCR of scanned pages.
- Changing existing words inside a PDF. Re-export from the source document instead.
- Creating PDFs with images, colour, several fonts, or non-Latin text.
- Digital signatures, redaction, extracting images, or converting Office files (use the docx or pptx skill).
- Removing encryption without the password.
