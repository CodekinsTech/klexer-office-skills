# XLSX skill (v2)

One command-line tool, `scripts/xlsx_tool.py`, works on `.xlsx` workbooks. Run it from any folder with an absolute or relative path to the file:

```
python "C:\path\to\xlsx\scripts\xlsx_tool.py" COMMAND FILE [options]
```

Rules that apply to every command:

- The tool never asks questions. A missing required option exits 1 and the message names the option.
- Every read command accepts `--json`. With `--json`, stdout is one JSON object. Without it, stdout is plain text. Errors and warnings always go to stderr.
- The input file is never overwritten, unless you pass `--in-place`.
- New files are written next to the input as `NAME.write.xlsx`, `NAME.format.xlsx` or `NAME.recalc.xlsx`, unless you pass `--out PATH`. An existing output is not replaced unless you pass `--force`.
- Every written file is re-opened and checked before it is moved into place. A failed check keeps nothing and exits 1.
- Cell references use A1 notation (`B2`), ranges use `A1:C10`. Sheet names are case-sensitive.
- Dates in JSON are ISO strings: `2026-10-24`, or `2026-10-24T14:30:00` when a time is present.
- Error values such as `#DIV/0!` are returned as text.

## Exit codes

- `0` success.
- `1` the input is missing or unreadable, a sheet, cell, range, patch or rules file is wrong, the output exists, a check found problems, or an internal error happened.
- `2` the file type is refused: legacy `.xls`, macro-enabled `.xlsm`, `.csv`/`.tsv`, or a password-protected workbook. Also exit 2 when a workbook has charts or drawing shapes and `--allow-drawing-rewrite` is not given (for `write`, `format` and `recalc`).

## Commands

### sheets
Lists the sheets.

- Options: `--json`.
- Output: one row per sheet with `name`, `state` (`visible`, `hidden` or `veryHidden`), `dimensions` (for example `A1:D20`), `max_row`, `max_column`. JSON: `{"sheets": [...]}`.
- Exit: 0 listed; 1 file missing or corrupt; 2 refused file type.

Example: `python xlsx_tool.py sheets book.xlsx --json`

### read
Reads cells or a range.

- Options: `--sheet NAME` (default: first visible sheet), `--range A1:C10` (default: the used range), `--mode values|formulas` (default `values`), `--out FILE` (also writes the JSON result to a new file), `--force`, `--json`.
- `values` mode returns the stored results from the last calculation. A formula with no stored result returns `null` and its cell appears in `not_calculated`. `formulas` mode returns formula text such as `=SUM(A1:A3)` for formula cells and plain values otherwise; `not_calculated` is empty in this mode.
- Output JSON: `{"sheet": "Data", "range": "A1:B1", "rows": [[10, null]], "not_calculated": ["B1"]}`. Text output is tab-separated rows.
- Exit: 0 read; 1 sheet or range not found, malformed range, or file error; 2 refused file type.

Example: `python xlsx_tool.py read book.xlsx --sheet Data --range A1:B1 --mode formulas --json`

### write
Writes values and formulas from a patch file.

- Options: `--patch FILE.json` (required), `--out FILE`, `--force`, `--in-place`, `--allow-drawing-rewrite`, `--json`.
- Patch file: a JSON list. Each entry is an object with `"sheet"` and `"cell"`, and exactly one of:
  - `"value"`: a JSON number is written as a number, a JSON string as text, `true`/`false` as booleans, `null` clears the cell.
  - `"formula"`: text that starts with `=`.
  - Optional `"type"`: `"text"`, `"number"` (numeric text is converted) or `"date"` (value must be `YYYY-MM-DD`; the cell gets the format `yyyy-mm-dd` unless `number_format` is given).
  - Optional `"number_format"`: an Excel number format code.
  - Any other key is refused.
- Formulas: functions added after Excel 2007 (for example `XLOOKUP`, `MAXIFS`, `MINIFS`, `TEXTJOIN`, `IFS`, `IFNA`) get the `_xlfn.` prefix that Excel needs; `FILTER` and `SORT` get `_xlfn._xlws.`; LET parameter names get `_xlpm.`. The output lists each one in `prefixed_functions`. A formula written here is one cell; a multi-value result spreads only after `recalc`.
- Output JSON: `{"output", "cells_written", "formula_cells", "prefixed_functions", "drawings_rewritten", "note"}`. `formula_cells` counts the formula cells in the whole workbook. Saving drops their stored results, so run `recalc` afterwards.
- Verification: the output is re-opened and each written cell must read back the given value or formula text.
- Exit: 0 written and verified; 1 patch malformed, sheet or cell not found, bad date, formula without `=`, output exists, verification failed, or file locked; 2 drawing guard or refused file type.

Example patch (`patch.json`):
```json
[{"sheet": "Data", "cell": "C1", "formula": "=XLOOKUP(20,A1:A3,A1:A3)"},
 {"sheet": "Data", "cell": "D1", "value": "2026-10-24", "type": "date"}]
```
Run: `python xlsx_tool.py write book.xlsx --patch patch.json --json`

### format
Applies a number format to a range.

- Options: `--sheet NAME`, `--range A1:C10` (default: used range), exactly one of `--preset currency|date|percent` or `--code "CODE"`, `--currency-symbol` (default `$`, used with the currency preset), `--out`, `--force`, `--in-place`, `--allow-drawing-rewrite`, `--json`.
- Presets: `currency` is `"$"#,##0.00` (the symbol is substituted), `date` is `yyyy-mm-dd`, `percent` is `0.00%`.
- Output JSON: `{"output", "sheet", "range", "number_format", "cells_formatted", "formula_cells", "drawings_rewritten", "note"}`.
- Exit: 0 formatted and verified; 1 range or sheet invalid, or both or neither of preset and code given; 2 drawing guard or refused file type.

Example: `python xlsx_tool.py format book.xlsx --sheet Data --range A1:A3 --preset currency`

### recalc
Calculates every formula and stores the results in a new file.

- Options: `--sheet NAME` (limits the JSON to one sheet; all sheets are still calculated), `--cells LIST` (comma-separated cells or ranges such as `B1,C1:C3` to include in the JSON), `--out`, `--force`, `--in-place`, `--timeout SECONDS` (default 180), `--allow-drawing-rewrite`, `--json`.
- Calculation is done by the tool itself. It supports these functions: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, IFERROR, SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIFS, VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, MAXIFS, MINIFS, TEXTJOIN, FILTER, SORT, UNIQUE, SEQUENCE and LET. Operators `+ - * / ^ & % = <> < > <= >=` and cell or range references (including other sheets) are supported; an operator applied to a range or array works cell by cell. Defined names work when they refer to one cell or range (workbook or sheet scope); a name that is not defined anywhere gives `#NAME?`; a name for a constant or formula is refused. XLOOKUP supports match modes 0, -1, 1, 2 and search modes 1, -1, 2, -2. A return range with several columns (or rows) gives the matching row (or column), which spills. Binary search modes (2, -2) and approximate lookups refuse unsorted data, and refuse a chosen value that occurs more than once.
- Array results spill. A formula whose result is more than one value (FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP returning a row or column, or LET returning an array) writes its values into the cells below and to the right of the formula cell. The file stores the spread as a dynamic array, and Excel 365 recognises it as a spill range (checked in Excel). The tool does not resize it when the source data changes: run `recalc` again. Blank results show as 0. If the spread would land on a non-empty cell, or on the spread of an earlier formula, the formula shows `#SPILL!` and nothing else is written. A single value is an ordinary result and does not spill. `#SPILL!` and `#CALC!` are stored as text in the file, because the file format has no code for them; Excel shows them as errors after it recalculates. A single-value formula that needs array evaluation (for example `=SUM(A1:A3*2)`) is saved as a one-cell array formula, so Excel gives the same result when it opens the file.
- LET names are local to their own LET call. Names that look like a cell (for example `A1`, `ABC1`, `R1C1`) and names such as `R1`, `C1` and `RC1` are refused. `R`, `C` and `RC` are accepted, as Excel accepts them.
- Refused, exit 1, nothing written. Exact messages (the cell and reason are filled in):
  - `LAMBDA`, including the immediately invoked form `LAMBDA(x,x*2)(5)`: `Formula in L!B1 is not supported by the built-in evaluator: the function LAMBDA is not supported. Nothing was written.`
  - Approximate or binary lookup on unsorted data: `Formula in L!B1 is not supported by the built-in evaluator: the lookup values are not sorted in the order this search needs (binary search on unsorted data is not implemented). Nothing was written.`
  - Any other function not listed above is refused the same way, with its name.
- If any formula uses something not supported, recalc exits 1, names the formula cell and the function, and writes nothing. Nothing partial is kept.
- Formula errors are not a failure of recalc: they are stored and listed in `errors`, and the exit code is still 0. Use `errors` to gate.
- Output JSON: `{"output", "sheet", "computed": {"Data!B1": 60}, "errors": {"Data!C1": "#DIV/0!"}, "drawings_rewritten"}`.
- Verification: every formula cell in the new file must have a stored result that matches the calculation, and every formula text must be unchanged.
- Exit: 0 recalculated and verified (errors inside formulas included); 1 timeout, unsupported function or syntax, circular reference, file error, or output problem.

Example: `python xlsx_tool.py recalc book.xlsx --sheet Data --cells B1 --json`

### errors
Lists formula errors and formulas that have no stored result.

- Options: `--sheet NAME` (default: all sheets), `--json`.
- Error values: `#DIV/0!`, `#N/A`, `#NAME?`, `#NULL!`, `#NUM!`, `#REF!`, `#VALUE!`, `#SPILL!`, `#CALC!`.
- Output JSON: `{"error_count": N, "not_calculated_count": N, "findings": [{"sheet", "cell", "formula", "value"}]}`. A formula with no stored result has `"value": null`.
- Exit: 0 no errors and no uncalculated formulas; 1 any finding; 2 refused file type.

Example: `python xlsx_tool.py errors book.recalc.xlsx --json`

### validate
Checks that the file is a valid workbook.

- Options: `--json`.
- Checks in order: readable zip; `[Content_Types].xml` and `xl/workbook.xml` present; every XML and `.rels` part parses; every internal relationship target exists; openpyxl opens the file; at least one sheet; sheet names are unique, at most 31 characters, and have none of `\ / ? * [ ] :`. The first failing check stops the run; the later checks are reported as `not run`.
- Output JSON: `{"file", "checks": [{"check", "result": "pass|fail|not run", "detail"}], "verdict": "valid|invalid"}`.
- Exit: 0 valid; 1 invalid or file missing; 2 refused file type (not checked).

Example: `python xlsx_tool.py validate book.xlsx --json`

### check-values
Checks test data in a workbook against a rules file. Nothing is written.

- Options: `--rules FILE.json` (required), `--sheet NAME` (default: first visible sheet), `--header-row N` (default 1), `--json`.
- Rules file: `{"columns": {"Header text": {rule, ...}}}`. The key is the header text in the header row. Rule keys:
  - `type` (required): `integer`, `decimal`, `date`, `boolean`, `text`. Numeric types need stored numbers. `date` needs stored dates, so text such as `"2026-02-30"` fails. `text` needs stored text.
  - `min`, `max`: numbers for integer and decimal columns, `YYYY-MM-DD` for date columns.
  - `allowed`: a list of permitted values.
  - `pattern`: a Python regular expression that must match the whole cell text.
  - `required` (default false): an empty cell fails.
  - `unique` (default false): repeated values fail, case-sensitive.
  - `max_length`: maximum number of characters.
  - Any other key makes the rules file invalid.
- Empty cells pass unless `required` is true. Fully empty rows are skipped and not counted in `rows_checked`. A formula with no stored result is reported with rule `not_calculated`.
- Output JSON: `{"sheet", "rows_checked": N, "violations": [{"sheet", "cell", "column", "rule", "value"}]}`. The `rule` value is one of `required`, `type`, `min`, `max`, `allowed`, `pattern`, `max_length`, `unique`, `not_calculated`.
- Exit: 0 no violations; 1 violations found, rules file invalid, or a named column missing from the header row; 2 refused file type.

Example rules file:
```json
{"columns": {"Name": {"type": "text", "required": true},
             "Age": {"type": "integer", "min": 0, "max": 120},
             "Start": {"type": "date"},
             "Code": {"type": "text", "pattern": "[A-Z][0-9]"}}}
```
Run: `python xlsx_tool.py check-values people.xlsx --rules rules.json --json`

## Common workflow

1. `write` (or `format`) to change the file. The output is a new `.write.xlsx`.
2. `recalc` on that new file to store formula results.
3. `errors` to confirm there are no errors and nothing is uncalculated. Exit 1 means fix the formulas first.
4. `read --mode values` to take the numbers.
