# PPTX v2 accuracy benchmark results

- Run: 2026-10-25 02:22
- Tool: `<skill-folder>\pptx\scripts\pptx_tool.py`
- Python: `<your-venv-folder>\venv-v2-pptx\Scripts\python.exe` (python-pptx 1.0.2)
- Tasks: 30, runnable: 30
- **Score: 30/30 passed (100%)**

## Score by category

| Category | Passed / runnable | Failed tasks |
|---|---|---|
| add | 2/2 | none |
| chart-data | 3/3 | none |
| check-text | 2/2 | none |
| create | 3/3 | none |
| delete | 1/1 | none |
| duplicate | 2/2 | none |
| read | 3/3 | none |
| refusal | 5/5 | none |
| reorder | 1/1 | none |
| replace | 3/3 | none |
| replace-image | 1/1 | none |
| split | 2/2 | none |
| validate | 2/2 | none |

## Top failure categories

- none

## Per-task results

| # | Task | Category | Result | Exit (expected / actual) |
|---|---|---|---|---|
| 1 | `A1_add_slide_levels_notes` Add a slide in the middle with sub-bullets and a note | add | PASS | 0 / 0 |
| 2 | `A2_add_position_out_of_range` Add at a position past the end is refused | add | PASS | 1 / 1 |
| 3 | `C1_create_outline_spec` Outline with levels, a speaker note and a second slide | create | PASS | 0 / 0 |
| 4 | `C2_create_text_before_title` Outline with a bullet before the first title is refused | create | PASS | 1 / 1 |
| 5 | `C3_create_unicode_roundtrip` Thai, Arabic and accented text survives create | create | PASS | 0 / 0 |
| 6 | `CD1_chart_data_update_keeps_type_and_colour` Replace a column chart's data; type and series colour stay | chart-data | PASS | 0 / 0 |
| 7 | `CD2_chart_data_count_mismatch` Chart data with more values than categories is refused | chart-data | PASS | 1 / 1 |
| 8 | `CD3_chart_data_xy_refused` Editing an XY scatter chart is refused | chart-data | PASS | 2 / 2 |
| 9 | `D1_delete_slide_removes_part` Delete slide 2 and check its content is gone from the file | delete | PASS | 0 / 0 |
| 10 | `I1_replace_image_keeps_box_and_name` Swap the one picture on a slide for a new image | replace-image | PASS | 0 / 0 |
| 11 | `K1_check_text_missing_phrase` Check phrases; one is missing and one is in the notes | check-text | PASS | 1 / 1 |
| 12 | `K2_check_text_whitespace_and_chart` Whitespace is normalised; chart categories and series names count | check-text | PASS | 0 / 0 |
| 13 | `O1_reorder_permutation` Move slide 3 to the front | reorder | PASS | 0 / 0 |
| 14 | `P1_replace_title_table_split_run_notes` Replace across title, table cell, split-run text and notes | replace | PASS | 0 / 0 |
| 15 | `P3_replace_pairs_file` Several replacements from a pairs file, notes included | replace | PASS | 0 / 0 |
| 16 | `R1_read_column_chart` Read a column chart's categories, series and values | read | PASS | 0 / 0 |
| 17 | `R2_read_table_and_slide_subset` Read one slide's table with --slides | read | PASS | 0 / 0 |
| 18 | `R3_read_never_creates_notes` Reading a deck does not change it or create a notes page | read | PASS | 0 / 0 |
| 19 | `S1_split_every_two` Split a five-slide deck into parts of two | split | PASS | 0 / 0 |
| 20 | `S2_split_uncovered_slide` Split ranges that leave a slide out are refused | split | PASS | 1 / 1 |
| 21 | `U1_duplicate_text_notes_layout` Duplicate slide 1 with bullets and notes | duplicate | PASS | 0 / 0 |
| 22 | `U2_duplicate_chart_refused` Duplicating a slide with a chart is refused and names the chart | duplicate | PASS | 2 / 2 |
| 23 | `V1_validate_valid_deck` A deck with notes and a picture passes every check | validate | PASS | 0 / 0 |
| 24 | `V2_validate_truncated_zip` A deck cut off at 1000 bytes fails validation | validate | PASS | 1 / 1 |
| 25 | `X1_refusal_corrupt_file` A corrupt .pptx is refused cleanly | refusal | PASS | 1 / 1 |
| 26 | `X2_refusal_missing_file` A missing input file is refused | refusal | PASS | 1 / 1 |
| 27 | `X3_refusal_legacy_ppt` A legacy .ppt is refused with exit 2 | refusal | PASS | 2 / 2 |
| 28 | `X4_refusal_macro_pptm` A macro-enabled .pptm is refused and not stripped | refusal | PASS | 2 / 2 |
| 29 | `X5_refusal_out_of_scope_render` A request to render slides to PDF is refused with exit 2 | refusal | PASS | 2 / 2 |
| 30 | `X7_replace_default_output_name` Default output is a new file; the input is never overwritten | replace | PASS | 0 / 0 |

## Failures

No failures.

