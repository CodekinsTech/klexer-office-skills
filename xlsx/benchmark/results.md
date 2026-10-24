# XLSX v2 skill accuracy benchmark: results

Run started: 2026-10-25 02:19:54
CLI: `<skill-folder>\xlsx\scripts\xlsx_tool.py`
Python: `<your-venv-folder>\venv-v2-xlsx\Scripts\python.exe`

**Score: 29/29 runnable tasks passed (100.0%).**
1 task(s) could not be set up and are not in the score (see the end).

## Score by category

| Category | Passed | Runnable | Score |
|---|---:|---:|---:|
| sheets | 1 | 1 | 100% |
| read | 3 | 3 | 100% |
| write | 5 | 5 | 100% |
| format | 2 | 2 | 100% |
| recalc | 8 | 8 | 100% |
| refusal | 4 | 4 | 100% |
| errors | 2 | 2 | 100% |
| validate | 1 | 1 | 100% |
| check-values | 3 | 3 | 100% |

No failures.

## Per-task results

| # | Task | Category | Result | Failed step(s) |
|---:|---|---|---|---|
| 1 | `sheets_lists_visible_and_hidden` | sheets | PASS | - |
| 2 | `read_values_uncalculated_formula` | read | PASS | - |
| 3 | `read_formulas_mode` | read | PASS | - |
| 4 | `read_cached_values_saved_file` | read | PASS | - |
| 5 | `refuse_missing_and_unsupported_files` | refusal | PASS | - |
| 6 | `write_xlookup_xlfn_prefix` | write | PASS | - |
| 7 | `write_date_cell` | write | PASS | - |
| 8 | `write_value_types_and_clear` | write | PASS | - |
| 9 | `write_formula_missing_equals` | write | PASS | - |
| 10 | `write_invalid_date` | write | PASS | - |
| 11 | `format_currency_preset` | format | PASS | - |
| 12 | `format_both_preset_and_code_refused` | format | PASS | - |
| 13 | `recalc_sum_and_arithmetic` | recalc | PASS | - |
| 14 | `recalc_lookup_functions` | recalc | PASS | - |
| 15 | `recalc_sumifs_family` | recalc | PASS | - |
| 16 | `recalc_let_and_iferror` | recalc | PASS | - |
| 17 | `recalc_filter_sort_unique_sequence` | recalc | PASS | - |
| 18 | `recalc_defined_name` | recalc | PASS | - |
| 19 | `recalc_cross_sheet_textjoin_maxifs` | recalc | PASS | - |
| 20 | `recalc_spill_blocked_shows_spill_error` | recalc | PASS | - |
| 21 | `refuse_lambda_on_recalc` | refusal | PASS | - |
| 22 | `refuse_unsorted_approximate_lookup` | refusal | PASS | - |
| 23 | `errors_div0_after_recalc` | errors | PASS | - |
| 24 | `errors_not_calculated_formula` | errors | PASS | - |
| 25 | `validate_valid_and_truncated` | validate | PASS | - |
| 26 | `check_values_spec_example` | check-values | PASS | - |
| 27 | `check_values_unique_allowed_maxlength` | check-values | PASS | - |
| 28 | `check_values_missing_column_refused` | check-values | PASS | - |
| 29 | `refuse_drawing_guard_then_allow` | refusal | PASS | - |

## Not runnable (excluded from the score)

- `refuse_password_protected`: ImportError: msoffcrypto-tool is not installed in the benchmark venv, so no password-protected file can be made

