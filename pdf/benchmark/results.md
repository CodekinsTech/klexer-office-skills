# PDF skill v2 benchmark results

- Run: 2026-10-25 01:47 (local)
- Python: 3.11.9 at `<your-venv-folder>\venv-v2-pdf\Scripts\python.exe`
- Platform: Windows-10-10.0.26200-SP0
- Tool under test: `<skill-folder>\pdf\scripts\pdf_tool.py`
- Tasks: 30 | Runnable: 30 | Not runnable: 0
- **Overall score: 30/30 (100%)**

## Score by category

| Group | Passed | Runnable | Score |
|---|---|---|---|
| text | 3 | 3 | 100% |
| password | 3 | 3 | 100% |
| tables | 1 | 1 | 100% |
| fields | 4 | 4 | 100% |
| merge | 2 | 2 | 100% |
| split | 2 | 2 | 100% |
| rotate | 4 | 4 | 100% |
| reorder | 2 | 2 | 100% |
| watermark | 2 | 2 | 100% |
| create | 2 | 2 | 100% |
| check | 4 | 4 | 100% |
| check-text | 1 | 1 | 100% |

## Score by task kind (happy path, refusal, safety)

| Group | Passed | Runnable | Score |
|---|---|---|---|
| happy | 16 | 16 | 100% |
| refusal | 8 | 8 | 100% |
| safety | 6 | 6 | 100% |

## Per-task results

| ID | Category | Kind | Result | Failed checks |
|---|---|---|---|---|
| 01_text_all_pages_json | text | happy | PASS | 0 of 10 checks failed |
| 02_text_scanned_refused | text | refusal | PASS | 0 of 3 checks failed |
| 03_text_page_list_out_of_range | text | happy | PASS | 0 of 3 checks failed |
| 04_password_no_env_refused | password | safety | PASS | 0 of 4 checks failed |
| 05_password_correct_env | password | happy | PASS | 0 of 4 checks failed |
| 06_password_wrong_no_leak | password | safety | PASS | 0 of 3 checks failed |
| 07_tables_invoice_rows | tables | happy | PASS | 0 of 5 checks failed |
| 08_fields_list_form | fields | happy | PASS | 0 of 5 checks failed |
| 09_fields_fill_ok | fields | happy | PASS | 0 of 8 checks failed |
| 10_fields_fill_unknown_no_write | fields | refusal | PASS | 0 of 4 checks failed |
| 11_fields_fill_no_form | fields | refusal | PASS | 0 of 3 checks failed |
| 12_merge_two_pdfs | merge | happy | PASS | 0 of 6 checks failed |
| 13_merge_output_exists_no_force | merge | safety | PASS | 0 of 3 checks failed |
| 14_split_every_2 | split | happy | PASS | 0 of 10 checks failed |
| 15_split_ranges_gap | split | refusal | PASS | 0 of 3 checks failed |
| 16_rotate_page_90 | rotate | happy | PASS | 0 of 8 checks failed |
| 17_rotate_wraps_360 | rotate | happy | PASS | 0 of 4 checks failed |
| 18_rotate_input_overwrite_refused | rotate | safety | PASS | 0 of 3 checks failed |
| 19_rotate_missing_angle | rotate | refusal | PASS | 0 of 3 checks failed |
| 20_reorder_3_1_2 | reorder | happy | PASS | 0 of 6 checks failed |
| 21_reorder_not_permutation | reorder | refusal | PASS | 0 of 3 checks failed |
| 22_watermark_stamp | watermark | happy | PASS | 0 of 5 checks failed |
| 23_watermark_text_draft | watermark | happy | PASS | 0 of 6 checks failed |
| 24_create_basic | create | happy | PASS | 0 of 4 checks failed |
| 25_create_unsupported_char | create | refusal | PASS | 0 of 4 checks failed |
| 26_check_valid | check | happy | PASS | 0 of 5 checks failed |
| 27_check_truncated_corrupt | check | safety | PASS | 0 of 3 checks failed |
| 28_check_empty_file | check | safety | PASS | 0 of 3 checks failed |
| 29_check_encrypted_no_pw | check | refusal | PASS | 0 of 3 checks failed |
| 30_checktext_missing_and_case | check-text | happy | PASS | 0 of 4 checks failed |

## Failures

No failures.

