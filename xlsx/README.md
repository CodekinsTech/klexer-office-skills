# XLSX skill (v2): developer notes

This folder holds the clean-room implementation of the xlsx skill from `specs\xlsx.md`. User-facing instructions for AI agents are in `SKILL.md`. This file covers libraries, licences, how to test, the decisions made where the spec left a detail open, and what is not done.

## Layout

- `SKILL.md`: command reference for agents (inputs, outputs, exit codes, examples, refused formulas with their exact messages).
- `scripts\xlsx_tool.py`: the single command-line tool, with all eight subcommands and the formula evaluator.
- `requirements.txt`: pinned runtime and test packages.
- `tests\test_acceptance.py`: one pytest test per acceptance test in spec section 3 (15 tests).
- `tests\test_formulas.py`: hand-computed tests for defined names, approximate MATCH and VLOOKUP, XLOOKUP match and search modes, and the refusals (11 tests).
- `tests\test_dynamic_arrays.py`: hand-computed tests for FILTER, SORT, UNIQUE, SEQUENCE and LET, spilling, the `#SPILL!` rules, LAMBDA refusal (plain and immediately invoked) and the `_xlfn` prefixes written by `write` (37 tests).
- `tests\test_excel_com.py`: compares `recalc` with Microsoft Excel through COM automation (10 tests: one per group of formulas, and a summary). Skipped automatically when Excel or pywin32 is not available, or when not on Windows.

## Libraries and licences

| Library | Used for | Licence (from installed metadata) |
|---|---|---|
| openpyxl 3.1.5 | reading and writing the cell model, number formats, charts and pictures | MIT |
| lxml 6.1.3 | parsing XML parts and relationship files in `validate`; writing cached values in `recalc` | BSD-3-Clause |
| Python standard library (`zipfile`, `json`, `re`, `tempfile`, `os.replace`, `argparse`, `datetime`, `posixpath`) | everything else | PSF licence |
| pytest 9.1.1 | tests only | MIT |
| pywin32 | Excel COM tests only (`test_excel_com.py`) | PSF licence |
| Pillow (optional, not installed) | keeps pictures on save; openpyxl uses it for images | MIT-CMU (HPND) |

No calculation engine is imported. The formula evaluator is code written for this skill (`Evaluator` and its parser in `xlsx_tool.py`). LibreOffice is not used and is not installed on this machine. No GPL or EUPL code was used. Microsoft Excel is used only as a reference in the tests; the tool does not call it.

## How to run the tests

From the `xlsx` folder, with the venv created for this task:

```
<your-venv-folder>\venv-v2-xlsx\Scripts\python.exe -m pytest tests -q
```

The Excel comparison needs pywin32 as well. Use the venv with pywin32 installed (the one made for the Excel tests):

```
<your-venv-folder>\venv-v2-xlsx-com\Scripts\python.exe -m pytest tests -q -s
```

`-s` prints the Excel comparison totals (formulas checked, formulas that differ, spill ranges, spilled cells, cells that differ).

Result at the time of writing: 73 passed, 0 failed (15 acceptance tests, 11 formula tests, 37 dynamic-array tests and 10 Excel comparison tests).

The tests build every input workbook with openpyxl, run the tool as a subprocess, and compare against expected values written out by hand in the test (the formula tests give the working in comments). The results are read back from the output file with openpyxl, not with the evaluator. The evaluator is never used to compute expected values. The Excel comparison uses Excel's own results as the expected values.

## Excel comparison (`tests\test_excel_com.py`)

- Excel: Microsoft 365, version 16.0, build 20430, started through COM with `DispatchEx`, not visible, alerts off. It is quit in a `finally` block.
- Each group is a workbook with a `Data` sheet, a `Calc` sheet with the formulas, and three defined names. The formulas are written the way Excel saves them in a file (`_xlfn.`, `_xlfn._xlws.`, `_xlpm.` prefixes).
- Excel calculates a copy of the workbook. Each formula is entered as `Formula2`, which is how a user in Microsoft 365 enters a formula that can spill. The prefixes are removed for that entry, as a user types the plain function name.
- The tool's `recalc` runs on the same workbook. The output is read back with openpyxl and compared with Excel: every formula value, every spilled cell, and every spill range (Excel's address against the array block the tool writes).
- The test also checks that Excel saves the same formula text as the input (prefix check), that `write` stores the same text Excel saves (so the prefixes `write` adds are checked), and that Excel opens the tool's output unchanged, calculates it, and gets the same values and spill ranges.
- Result on the final run: 156 formulas, 28 spill ranges with 213 spilled cells. 0 differ. Differences found during the work, and how each was fixed, are in decisions 30 to 36.

## Formula support (`recalc`)

Supported functions: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, IFERROR, SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIFS, MAXIFS, MINIFS, TEXTJOIN, VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE, LET.

Supported features: operators `+ - * / ^ & % = <> < > <= >=`, applied cell by cell when an operand is a range or array (both sides the same size, or one side a single value); cell and range references, including other sheets; defined names that refer to one cell or range (workbook scope or sheet scope); TRUE and FALSE.

Lookup modes:
- MATCH: match type 0 (linear), 1 and -1 (binary search on sorted data).
- VLOOKUP and HLOOKUP: FALSE (linear), TRUE (binary search on sorted data).
- XLOOKUP: match modes 0, -1, 1 and 2 with search modes 1 and -1 (linear); match modes 0, -1 and 1 with search modes 2 and -2 (binary search on sorted data). With match modes -1 and 1, an exact match is taken first. A return range with several columns (vertical lookup) or rows (horizontal lookup) gives the matching row or column.

Refused with exit 1 and a message naming the cell: `LAMBDA` (plain or immediately invoked, as `LAMBDA(x,x*2)(5)`) and every other function not listed above (for example `LAMBDA(x,x+1)` gives `the function LAMBDA is not supported`); a multi-cell array formula from the source file that is not one of the array functions; defined names that refer to a constant or a formula; `Sheet!Name` (a sheet-scope name used with its sheet prefix); MATCH match types other than 0, 1 and -1; wildcard matching with a binary search mode; approximate or binary lookups on unsorted data; approximate or binary lookups whose chosen value occurs more than once; LET names that look like cell references (decision 26); circular references. The exact messages are listed in `SKILL.md`.

A defined name that does not exist anywhere in the workbook gives `#NAME?` as a stored result, as Excel does.

## Dynamic arrays and spilling

- `FILTER`, `SORT`, `UNIQUE` and `SEQUENCE` return arrays; `LET` returns whatever its last calculation returns; `XLOOKUP` returns a row or column when its return range has several columns or rows. A result of one value is an ordinary cell value.
- A formula whose result has more than one value writes the values into the cells below and to the right of the formula cell. The result is stored as an array formula over its block (`t="array"` with a `ref`), on a cell marked as dynamic (`cm="1"`). The workbook gets the part `xl/metadata.xml` that Excel uses for that mark (decision 31). The other block cells hold the cached values, as Excel stores them. Excel 365 reads the block as a spill range.
- A block may not cover a non-empty cell (a constant or another formula), run past the last row or column, or overlap the block of a formula calculated earlier. In those cases the formula shows `#SPILL!` and nothing else is written. Formulas are calculated row by row from the top left, so the earlier one keeps its block.
- Blank results are written as 0, as Excel shows a reference to a blank cell in a result.
- A formula that reads a cell inside a block of a formula it depends on is a circular reference and is refused.
- Blocks are recalculated, not kept: old block cells are emptied before a new run, so a block that shrinks loses its extra cells.
- `write` and `format` empty the block cells on save, since saving drops stored results, and openpyxl does not keep the dynamic mark. Run `recalc` afterwards.

## Decisions where the spec was open

1. **Recalculation engine.** Spec section 7 left the engine open (LibreOffice or a pure-Python package). The brief for this build says to write a small evaluator for the listed functions and not to use GPL or EUPL engines. The `recalc` command uses that evaluator. It needs no external program, so exit 2 ("no calculation engine") is never produced.
2. **Unsupported formulas.** When a formula uses something the evaluator does not implement (any function not on the list, `LAMBDA`, a defined name that does not refer to one cell or range, a multi-cell result in a single-value position, `INDEX` returning a range, a binary search on unsorted data or on a repeated chosen value, a wildcard with a binary search mode, `MATCH` match types other than 0, 1 and -1, or a circular reference), `recalc` exits 1, names the cell and reason, and writes no file. Results are never guessed.
3. **Approximate match and binary search.** `VLOOKUP`/`HLOOKUP`/`MATCH` with approximate matching, and `XLOOKUP` search modes 2 and -2, use a binary search over the values of the same type as the key; blanks and values of another type are skipped. Data in the wrong order is refused as above. If the value the search chooses occurs more than once, the case is refused, because Excel's choice among equal values is not documented. The rules were taken from Microsoft's function descriptions and checked in Excel for the approximate and binary cases in `test_excel_com.py` (groups `approximate_lookups` and `xlookup_modes`). Repeated chosen values were not tried in Excel, so they stay refused.
4. **`_xlfn` prefix list.** The list in `scripts\xlsx_tool.py` (`XLFN_FUNCTIONS`) was written from memory of Excel's post-2007 function set. Confirmed in Excel (test_excel_com.py, the saved text Excel writes equals the text `write` stores): `_xlfn.` for XLOOKUP, MAXIFS, MINIFS, TEXTJOIN, UNIQUE, SEQUENCE and LET; `_xlfn._xlws.` for FILTER and SORT; `_xlpm.` for LET parameter names. The rest of the list (for example `IFS`, `IFNA`, `SWITCH`, `CONCAT`, the statistical `.DIST` functions) is still not confirmed against Excel, and those functions are refused by `recalc` anyway.
5. **Read output.** `not_calculated` is a top-level list of cell references, not a per-row flag. In `formulas` mode it is always empty. `read --out` always writes the JSON result, whatever `--json` says.
6. **Options added beyond spec section 2.** `--in-place` (named in the error table and safety rules) is on `write`, `format` and `recalc`. `--allow-drawing-rewrite` is on `format` and `recalc` as well as `write`, since all three re-save the file.
7. **Option defaults.** `format` and `read` without `--range` use the used range. Without `--sheet`, the first visible sheet is used. `recalc --cells` applies to `--sheet` if given, otherwise to the first visible sheet. `recalc --sheet` limits only the JSON; all sheets are calculated.
8. **Argument errors.** argparse's own usage errors are changed to exit 1, because the spec requires a missing option to exit 1.
9. **Patch file.** Unknown keys are refused. A value that is text starting with `=` is refused with a message telling the caller to use `formula`. `type: "text"` writes numbers and booleans as text (`TRUE`/`FALSE`). `type: "number"` converts numeric text. A `null` value clears the cell.
10. **Formula results after save.** openpyxl does not keep stored formula results when it re-saves. So `write` and `format` drop them for every formula in the workbook, not only the patched cells. Both commands say so in the output `note`, and `formula_cells` counts all formula cells in the workbook.
11. **Write verification.** Formula cells are checked against the prefixed formula text (`=_xlfn.XLOOKUP(...)`), not the text the caller typed.
12. **Check-values rules file.** `columns` is an object keyed by header text, one rule object per column. `min`/`max` apply to integer, decimal and date columns only. Numeric types need stored numbers, `date` needs stored dates (text dates are violations), `text` needs stored text. Whitespace-only text counts as empty. `unique` is case-sensitive. If a header appears twice, the first occurrence is used. A formula cell with no stored result is a violation with rule `not_calculated`. When a cell fails its type rule, its other rules are not checked.
13. **Validate.** Checks run in order and stop at the first failure; later checks show `not run`. Sheet-name uniqueness ignores case, as Excel does.
14. **Drawing guard.** A workbook is treated as having drawings when it has chart parts or non-picture drawing objects. Pictures count as drawings when Pillow is not installed, because openpyxl would drop them.
15. **Timeout.** `--timeout` is checked between formula evaluations, not enforced by killing a process. A single very large `SUMIFS` may run past the limit slightly.
16. **Errors and messages.** The spec's error messages are used as written. Some extra messages were added for cases the table does not list (for example "Output must be an .xlsx file", "Not an .xlsx file", "Patch file not found", "Output folder does not exist", "Circular reference at ...", "Recalculation timed out after N seconds.", "--timeout must be a positive number of seconds."). Result JSON goes to stdout and check-command summaries go to stderr.
17. **Encrypted files.** A file with the OLE header but a `.xlsx` name is reported as password-protected (exit 2). A legacy `.xls` file renamed to `.xlsx` would get the same message.
18. **Defined names.** A workbook-level or sheet-level name whose stored text is one cell, one range or a whole-column range (for example `Data!$A$1:$A$3`, `'My Data'!B2`) resolves to that range. A sheet-level name is looked up on the formula's own sheet first. A name that is defined nowhere gives `#NAME?`. A name that refers to a constant or a formula is refused, not evaluated.
19. **XLOOKUP linear modes.** Search modes 1 and -1 scan in the stated order and return the first match found, so with repeated values the first (or last, for -1) row wins, as the search mode describes.
20. **MATCH match_type.** Only 0, 1 and -1 are accepted. Other values were `#VALUE!` in the first build; they are now refused, because the result for them is not documented.
21. **FILTER include values.** TRUE and non-zero numbers keep a row; FALSE, zero and blank drop it; text gives `#VALUE!`; an error in the include values is returned. The include array must match the array's height (rows) or width (columns), or the result is `#VALUE!`. Checked in Excel: a numeric include keeps the non-zero rows (group `dynamic_arrays`). Text include values were not tried in Excel.
22. **FILTER with no match.** `#CALC!` when `if_empty` is omitted, as the public page says. `if_empty` is returned otherwise, as a block when it is a range. Checked in Excel: both forms.
23. **SORT.** `sort_order` must be 1 or -1, and `sort_index` must be inside the array; otherwise `#VALUE!`. Text sorts ignoring case; numbers come before text and TRUE/FALSE after it. Blanks sort last in both orders and read as 0 in the result. Checked in Excel: text and numeric sorts, ties keep their order (stable), and descending order. Blank order and an out-of-range `sort_index` were not tried in Excel.
24. **UNIQUE.** Compares rows (or columns with `by_col`) and keeps the first occurrence. Text is compared ignoring case. A blank is its own value. `exactly_once` keeps the values that occur once; when none does, the result is `#CALC!`. Checked in Excel: first occurrence, case-insensitive text, row uniqueness and `exactly_once`. The `#CALC!` for "none occurs once" was not tried in Excel.
25. **SEQUENCE.** `rows` and `columns` below 1 give `#CALC!`. Checked in Excel (`SEQUENCE(0)` gives `#CALC!`). A result over one million cells is refused.
26. **LET.** Names are case-insensitive, must start with a letter, underscore or backslash, and may be given only once per call. A name is visible only inside its own call; a name used outside it falls back to the defined names, then `#NAME?`. A LET name shadows a defined name inside the call. Name rules (checked in Excel 365): `R`, `C`, `RC` and similar names with no digits are accepted; names such as `R1`, `C1`, `RC1`, `R100` and `C5` are refused, as Excel refuses them; names that look like cell references (`A1`, `ABC1`) and names in R1C1 form with digits (`R1C1`, `R2C3`) are refused too. Excel accepts those in a LET call but gives `#NAME?` for cell references of the same shape in the same formula (for example `=LET(A1,A2,A1*2)`), so no value is guessed for them.
27. **Operators on arrays.** An operator with a range or array operand works cell by cell. Two arrays must be the same size. Any other use of an array where one value is needed (for example `IF(A1:A3>1, ...)`) is refused.
28. **Array arguments.** `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `TEXTJOIN`, `FILTER`, `SORT` and `UNIQUE` accept an array argument. Every other function accepts an array only when it has one element; otherwise the formula is refused.
29. **Overlapping blocks and #SPILL!.** The first formula calculated keeps its block; a later one that would overlap it, or cover a non-empty cell, shows `#SPILL!` and writes nothing else. Checked in Excel: a block blocked by a constant, and a block blocked by a formula cell in its range (`test_excel_com.py`, group `spill_blocking`).
30. **Formula results Excel gave that the first build did not match (fixed).** (a) XLOOKUP in binary mode with match mode -1 or 1 returned the neighbouring value when the lookup value was in the range (for example `XLOOKUP(4, E1:E8, E1:E8, "nf", -1, 2)` gave 3, Excel gives 4). The binary search now takes the exact match first. (b) XLOOKUP with a multi-column return range gave `#VALUE!`; Excel returns the matching row, which spills. It is now calculated and spilled. (c) `=SUM(A1:A8*2)`-style formulas and other single-value formulas that need array evaluation were calculated with array semantics, but a file opened in Excel calculated them with implicit intersection (`#VALUE!` for that formula). Such formulas are now saved as one-cell array formulas, so Excel gives the same value (decision 33).
31. **Dynamic-array metadata (fixed).** The first build wrote spills as plain array blocks without the dynamic mark. Excel then treated them as legacy arrays: a blocked or empty result opened in Excel differed from the tool. Dynamic formulas are now marked with `cm="1"` and the workbook gets `xl/metadata.xml` with one dynamic-array property, a minimal form of the part Excel writes (Excel's own saved file also has rich-value parts for error types, which this tool does not need). Excel 365 then reads the block as a spill range; the test checks that Excel reports the same spill addresses for the tool's file.
32. **Error values the file format lacks (fixed).** `#SPILL!` and `#CALC!` are not among the error codes in the file format. The first build wrote them as error cells, and Excel refused to open the file ("Open method of Workbooks class failed"). Excel's own saved files store them as `#VALUE!` with extra value metadata. The tool stores them as text (`t="str"`), so the file opens; `errors` still lists them, because it reads the stored text. Excel shows the real error after it recalculates (checked in Excel).
33. **Single-value formulas that need array evaluation.** An operator on a range or array, a FILTER, SORT, UNIQUE or SEQUENCE call, or an XLOOKUP returning a range, makes a formula an array formula. If its result is one value, it is saved as a one-cell array formula (`t="array"` with `ref` equal to the cell, and `cm="1"`). Its value is the same as the value Excel gives when the formula is entered as a dynamic array, for example `SUM(A1:A8*2)` = 80.
34. **Spill errors inside blocks.** A spill blocked by a cell, or an empty FILTER with no `if_empty`, is a single-value array formula stored with `#SPILL!` or `#CALC!` as text (decision 32). Excel gives the same error after it recalculates the file (checked).
35. **Excel version used for the comparison.** Microsoft 365 (version 16.0, build 20430). Formulas were entered as `Formula2`. Excel's behaviour for a formula typed as a legacy formula (`Formula`, which Excel 365 shows with `@`) was not used as the reference, because a user typing a FILTER, SORT, UNIQUE or SEQUENCE formula in Microsoft 365 gets the dynamic form.
36. **Not a bug: Excel's `@` for legacy formulas.** When Excel opens a file that has a plain formula (no dynamic mark) containing FILTER or SORT, Excel shows it with `@` and returns the first value. The tool never writes such a formula: every formula that needs array evaluation is written with the mark or as an array formula (decisions 31 and 33).

## What is not done

- Excel 365 (build 16.0.20430) was run through COM for the formulas in `tests\test_excel_com.py` only. LibreOffice and other engines were not run. The sheets and cells of the Excel test are generated; no real user workbook was checked.
- The `_xlfn.` list is confirmed only for the functions in decision 4. The rest is still open.
- The `#SPILL!` and `#CALC!` values are stored as text, not as Excel's own error type (decision 32). A user reading the raw file sees text until Excel recalculates it.
- `write` and `format` open the file with openpyxl, which drops the dynamic-array mark. After either command, run `recalc` to store the results and the marks again (decision 31).
- Not supported: `LAMBDA`, `SORTBY`, `RANDARRAY`, the `#` spill-range operator, and an array result passed to any function other than the ones named in decision 28 (for example to a lookup or `COUNTIF`). Each is refused with the message in `SKILL.md`.
- Not tested on macOS or Linux, and not tested with a workbook locked by Excel on Windows (the locked-file message is wired but not exercised).
- Pictures and the Pillow path are not covered by a test.
- Conditional formatting and data validation are kept by openpyxl but not created. Pivot tables, data models and drawing shapes are not created or edited (only the guarded re-save).
- Legacy `.xls`, `.xlsm` (not stripped), and password-protected files are refused, as the spec requires.
- The chart guard and the `--allow-drawing-rewrite` path are tested for the flag and the JSON field only. Whether the chart's formatting survives openpyxl's re-save is not compared.
- Git: nothing was committed.
