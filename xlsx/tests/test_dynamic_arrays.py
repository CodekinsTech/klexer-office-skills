"""Tests for FILTER, SORT, UNIQUE, SEQUENCE and LET, and for spilling array results.

Every expected value is written out by hand in the test, with the working in a comment, never taken
from the evaluator. Results are read back from the output file with openpyxl (data_only), so the
check does not depend on the tool's own reader. Refusals are checked for exit code 1, the message,
and that no output file is kept.
"""

import json
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "xlsx_tool.py"


def run(*args):
    return subprocess.run([sys.executable, str(SCRIPT), *[str(a) for a in args]],
                          capture_output=True, text=True, encoding="utf-8")


def make_book(path, columns, formulas, sheet="L"):
    """columns: {"A": [values from row 1 down]}; formulas: {"B1": "=..."}. No cached values."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for col, values in columns.items():
        for row, value in enumerate(values, start=1):
            ws[f"{col}{row}"] = value
    for cell, formula in formulas.items():
        ws[cell] = formula
    wb.save(path)
    return path


def recalc_ok(book):
    proc = run("recalc", book, "--json")
    assert proc.returncode == 0, proc.stderr
    return book.with_name(book.stem + ".recalc.xlsx")


def read_back(out_path, cells, sheet="L"):
    """Cached values of the given cells in an output file, read with openpyxl only."""
    wb = openpyxl.load_workbook(out_path, data_only=True)
    return {c: wb[sheet][c].value for c in cells}


def assert_refused(book, fragment):
    proc = run("recalc", book, "--json")
    assert proc.returncode == 1, proc.stdout + proc.stderr
    assert fragment in proc.stderr
    assert not book.with_name(book.stem + ".recalc.xlsx").exists()


# Shared data on sheet L:
#   A1:A5 = 10, 20, 30, 40, 50
#   B1:B5 = a, b, c, d, e
SALES = {"A": [10, 20, 30, 40, 50], "B": ["a", "b", "c", "d", "e"]}


# ---------------------------------------------------------------- FILTER

def test_filter_spills_matching_rows_below_the_formula(tmp_path):
    # rows with A > 25 are 30, 40, 50 (rows 3, 4 and 5), so they fill B1:B3
    book = make_book(tmp_path / "f1.xlsx", SALES, {"D1": "=FILTER(A1:A5,A1:A5>25)"})
    out = recalc_ok(book)
    assert read_back(out, ["D1", "D2", "D3", "D4"]) == {"D1": 30, "D2": 40, "D3": 50, "D4": None}


def test_filter_keeps_matching_columns_when_include_is_a_row(tmp_path):
    # rows: [1, 2, 3] and [4, 5, 6]; include row A3:C3 = 1, 0, 1 keeps columns A and C
    columns = {"A": [1, 4, 1], "B": [2, 5, 0], "C": [3, 6, 1]}
    book = make_book(tmp_path / "f2.xlsx", columns, {"E1": "=FILTER(A1:C2,A3:C3)"})
    out = recalc_ok(book)
    assert read_back(out, ["E1", "F1", "E2", "F2", "G1"]) == {
        "E1": 1, "F1": 3, "E2": 4, "F2": 6, "G1": None}


def test_filter_if_empty_is_returned_when_nothing_matches(tmp_path):
    book = make_book(tmp_path / "f3.xlsx", SALES, {"D1": '=FILTER(A1:A5,A1:A5>99,"none")'})
    out = recalc_ok(book)
    assert read_back(out, ["D1", "D2"]) == {"D1": "none", "D2": None}


def test_filter_single_result_is_a_plain_value(tmp_path):
    # only row 3 matches, so the result is one value and nothing spills
    book = make_book(tmp_path / "f4.xlsx", SALES, {"D1": "=FILTER(A1:A5,A1:A5=30)"})
    out = recalc_ok(book)
    assert read_back(out, ["D1", "D2"]) == {"D1": 30, "D2": None}


def test_filter_no_match_without_if_empty_is_calc_error(tmp_path):
    # Microsoft: #CALC! appears when no rows match and if_empty is omitted
    book = make_book(tmp_path / "f5.xlsx", SALES, {"D1": "=FILTER(A1:A5,A1:A5>99)"})
    out = recalc_ok(book)
    assert read_back(out, ["D1"]) == {"D1": "#CALC!"}


def test_filter_include_size_mismatch_is_value_error(tmp_path):
    # include has 3 rows but the array has 5, so no orientation fits: #VALUE!
    columns = dict(SALES, C=[1, 1, 1])
    book = make_book(tmp_path / "f6.xlsx", columns, {"D1": "=FILTER(A1:A5,C1:C3>0)"})
    out = recalc_ok(book)
    assert read_back(out, ["D1"]) == {"D1": "#VALUE!"}


def test_filter_text_in_include_is_value_error(tmp_path):
    # Microsoft: an include value that cannot be read as TRUE/FALSE gives an error
    columns = dict(SALES, D=["yes", "no", "yes", "no", "yes"])
    book = make_book(tmp_path / "f7.xlsx", columns, {"E1": "=FILTER(A1:A5,D1:D5)"})
    out = recalc_ok(book)
    assert read_back(out, ["E1"]) == {"E1": "#VALUE!"}


# ---------------------------------------------------------------- SORT

def test_sort_ascending_spills(tmp_path):
    # 30, 10, 40, 20 sorted ascending is 10, 20, 30, 40
    book = make_book(tmp_path / "s1.xlsx", {"A": [30, 10, 40, 20]}, {"B1": "=SORT(A1:A4)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3", "B4"]) == {"B1": 10, "B2": 20, "B3": 30, "B4": 40}


def test_sort_descending_by_second_column_keeps_all_columns(tmp_path):
    # rows (Ana,3), (Ben,1), (Cy,2); descending by column 2 gives Ana 3, Cy 2, Ben 1
    columns = {"A": ["Ana", "Ben", "Cy"], "B": [3, 1, 2]}
    book = make_book(tmp_path / "s2.xlsx", columns, {"D1": "=SORT(A1:B3,2,-1)"})
    out = recalc_ok(book)
    assert read_back(out, ["D1", "E1", "D2", "E2", "D3", "E3"]) == {
        "D1": "Ana", "E1": 3, "D2": "Cy", "E2": 2, "D3": "Ben", "E3": 1}


def test_sort_text_ignores_case(tmp_path):
    # b, C, a: ignoring case the order is a, b, C (a plain code-point sort would give C, a, b)
    book = make_book(tmp_path / "s3.xlsx", {"A": ["b", "C", "a"]}, {"B1": "=SORT(A1:A3)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3"]) == {"B1": "a", "B2": "b", "B3": "C"}


def test_sort_by_column_orders_the_columns(tmp_path):
    # columns (30,x), (10,y), (20,z); sorting by row 1 ascending gives (10,y), (20,z), (30,x)
    columns = {"A": [30, "x"], "B": [10, "y"], "C": [20, "z"]}
    book = make_book(tmp_path / "s4.xlsx", columns, {"D1": "=SORT(A1:C2,1,1,TRUE)"})
    out = recalc_ok(book)
    assert read_back(out, ["D1", "E1", "F1", "D2", "E2", "F2"]) == {
        "D1": 10, "E1": 20, "F1": 30, "D2": "y", "E2": "z", "F2": "x"}


def test_sort_blanks_go_last_in_both_orders(tmp_path):
    # my choice, not confirmed by Microsoft: blanks stay last; a blank in a result reads as 0
    book = make_book(tmp_path / "s5.xlsx", {"A": [2, None, 1]}, {"B1": "=SORT(A1:A3,1,-1)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3"]) == {"B1": 2, "B2": 1, "B3": 0}


def test_sort_order_other_than_one_or_minus_one_is_value_error(tmp_path):
    book = make_book(tmp_path / "s6.xlsx", {"A": [3, 1, 2]}, {"B1": "=SORT(A1:A3,1,2)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1"]) == {"B1": "#VALUE!"}


def test_sort_index_outside_the_array_is_value_error(tmp_path):
    # the array has one column, so sort_index 3 is out of range: #VALUE!
    book = make_book(tmp_path / "s7.xlsx", {"A": [3, 1, 2]}, {"B1": "=SORT(A1:A3,3)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1"]) == {"B1": "#VALUE!"}


# ---------------------------------------------------------------- UNIQUE

def test_unique_keeps_first_occurrence_ignoring_case(tmp_path):
    # Apple, pear, apple, fig, pear, kiwi -> Apple, pear, fig, kiwi
    columns = {"A": ["Apple", "pear", "apple", "fig", "pear", "kiwi"]}
    book = make_book(tmp_path / "u1.xlsx", columns, {"B1": "=UNIQUE(A1:A6)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3", "B4", "B5"]) == {
        "B1": "Apple", "B2": "pear", "B3": "fig", "B4": "kiwi", "B5": None}


def test_unique_exactly_once_keeps_single_values(tmp_path):
    # 1, 2, 2, 3: the values that occur once are 1 and 3
    book = make_book(tmp_path / "u2.xlsx", {"A": [1, 2, 2, 3]}, {"B1": "=UNIQUE(A1:A4,FALSE,TRUE)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3"]) == {"B1": 1, "B2": 3, "B3": None}


def test_unique_rows_of_two_columns(tmp_path):
    # rows (1,x), (1,x), (2,y), (1,z); the unique rows are (1,x), (2,y), (1,z)
    columns = {"A": [1, 1, 2, 1], "B": ["x", "x", "y", "z"]}
    book = make_book(tmp_path / "u3.xlsx", columns, {"D1": "=UNIQUE(A1:B4)"})
    out = recalc_ok(book)
    assert read_back(out, ["D1", "E1", "D2", "E2", "D3", "E3"]) == {
        "D1": 1, "E1": "x", "D2": 2, "E2": "y", "D3": 1, "E3": "z"}


def test_unique_exactly_once_with_no_single_value_is_calc_error(tmp_path):
    # 1, 1, 2, 2: no value occurs exactly once, so the result is empty: #CALC!
    book = make_book(tmp_path / "u4.xlsx", {"A": [1, 1, 2, 2]}, {"B1": "=UNIQUE(A1:A4,FALSE,TRUE)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1"]) == {"B1": "#CALC!"}


# ---------------------------------------------------------------- SEQUENCE

def test_sequence_single_column_spills(tmp_path):
    book = make_book(tmp_path / "q1.xlsx", {}, {"B1": "=SEQUENCE(3)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3", "B4"]) == {"B1": 1, "B2": 2, "B3": 3, "B4": None}


def test_sequence_rows_columns_start_and_step(tmp_path):
    # SEQUENCE(2,3,10,5): row 1 is 10, 15, 20; row 2 is 25, 30, 35
    book = make_book(tmp_path / "q2.xlsx", {}, {"B1": "=SEQUENCE(2,3,10,5)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "C1", "D1", "B2", "C2", "D2"]) == {
        "B1": 10, "C1": 15, "D1": 20, "B2": 25, "C2": 30, "D2": 35}


def test_sequence_with_zero_rows_is_calc_error(tmp_path):
    # the public page does not state this error; #CALC! is the chosen value (see README)
    book = make_book(tmp_path / "q3.xlsx", {}, {"B1": "=SEQUENCE(0)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1"]) == {"B1": "#CALC!"}


# ---------------------------------------------------------------- LET

def test_let_names_hold_values_used_by_later_names(tmp_path):
    # x = 2, y = x * 3 = 6, result x + y = 8
    book = make_book(tmp_path / "l1.xlsx", {}, {"B1": "=LET(x,2,y,x*3,x+y)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1"]) == {"B1": 8}


def test_let_name_hides_a_defined_name_only_inside_the_call(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "L"
    ws["A1"] = 5
    ws["B1"] = "=LET(total,7,total+1)"   # the LET name total is 7, so 8 (not 5 + 1)
    ws["C1"] = "=total"                  # outside LET the defined name applies: A1, so 5
    wb.defined_names["total"] = DefinedName("total", attr_text="L!$A$1")
    book = tmp_path / "l2.xlsx"
    wb.save(book)
    out = recalc_ok(book)
    assert read_back(out, ["B1", "C1"]) == {"B1": 8, "C1": 5}


def test_let_name_is_not_visible_in_another_cell(tmp_path):
    # x exists only inside the LET in B1; a plain x in B2 has no definition, so #NAME?
    book = make_book(tmp_path / "l3.xlsx", {}, {"B1": "=LET(x,1,x)", "B2": "=x"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2"]) == {"B1": 1, "B2": "#NAME?"}


def test_let_holds_a_range_and_spills_an_array(tmp_path):
    # A1:A3 = 1, 2, 3; SUM is 6; SORT descending is 3, 2, 1 in C1:C3
    book = make_book(tmp_path / "l4.xlsx", {"A": [1, 2, 3]},
                     {"B1": "=LET(v,A1:A3,SUM(v))", "C1": "=LET(v,A1:A3,SORT(v,1,-1))"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "C1", "C2", "C3"]) == {"B1": 6, "C1": 3, "C2": 2, "C3": 1}


def test_let_with_a_cell_reference_name_is_refused(tmp_path):
    book = make_book(tmp_path / "l5.xlsx", {}, {"B1": "=LET(A1,1,A1)"})
    assert_refused(book, '"A1" is not a valid LET name')


def test_let_without_a_calculation_is_refused(tmp_path):
    book = make_book(tmp_path / "l6.xlsx", {}, {"B1": "=LET(x,1,x,2)"})
    assert_refused(book, "LET needs one or more name and value pairs followed by a calculation")


# ---------------------------------------------------------------- refusals kept

def test_lambda_is_refused_with_its_name_in_the_message(tmp_path):
    book = make_book(tmp_path / "x1.xlsx", {}, {"B1": "=LAMBDA(x,x+1)"})
    assert_refused(book, "the function LAMBDA is not supported")


def test_immediately_invoked_lambda_gives_the_same_refusal(tmp_path):
    book = make_book(tmp_path / "x2.xlsx", {}, {"B1": "=LAMBDA(x,x*2)(5)"})
    assert_refused(book, "Formula in L!B1 is not supported by the built-in evaluator: "
                         "the function LAMBDA is not supported.")


def test_calling_a_non_function_value_still_says_unexpected_text(tmp_path):
    book = make_book(tmp_path / "x3.xlsx", {}, {"B1": "=SUM(1)(2)"})
    assert_refused(book, "unexpected text after the formula")


# ---------------------------------------------------------------- spilling rules

def test_spill_blocked_by_a_constant_gives_spill_error_and_writes_nothing_else(tmp_path):
    # B2 holds "keep", so the FILTER result cannot spill: B1 shows #SPILL!, B2 and B3 stay as they were
    book = make_book(tmp_path / "p1.xlsx", SALES,
                     {"D1": "=FILTER(A1:A5,A1:A5>25)"})
    wb = openpyxl.load_workbook(book)
    wb["L"]["D2"] = "keep"
    wb.save(book)
    out = recalc_ok(book)
    assert read_back(out, ["D1", "D2", "D3"]) == {"D1": "#SPILL!", "D2": "keep", "D3": None}


def test_spill_result_can_be_read_by_another_formula(tmp_path):
    # the spilled values are 30, 40, 50 in D1:D3, so SUM(D1:D3) is 120
    book = make_book(tmp_path / "p2.xlsx", SALES,
                     {"D1": "=FILTER(A1:A5,A1:A5>25)", "F1": "=SUM(D1:D3)"})
    out = recalc_ok(book)
    assert read_back(out, ["F1"]) == {"F1": 120}


def test_recalc_again_keeps_the_spill_and_shrinks_it_when_data_changes(tmp_path):
    book = make_book(tmp_path / "p3.xlsx", SALES, {"D1": "=FILTER(A1:A5,A1:A5>25)"})
    first = recalc_ok(book)
    assert read_back(first, ["D1", "D2", "D3"]) == {"D1": 30, "D2": 40, "D3": 50}
    # recalc on its own output: the spilled cells are owned by D1, so no #SPILL!
    again = tmp_path / "p3_again.xlsx"
    assert run("recalc", first, "--out", again, "--json").returncode == 0
    assert read_back(again, ["D1", "D2", "D3"]) == {"D1": 30, "D2": 40, "D3": 50}
    # change A5 to 1: only 30 and 40 pass, so D3 must be cleared
    wb = openpyxl.load_workbook(again)
    wb["L"]["A5"] = 1
    changed = tmp_path / "p3_changed.xlsx"
    wb.save(changed)
    shrunk = recalc_ok(changed)
    assert read_back(shrunk, ["D1", "D2", "D3"]) == {"D1": 30, "D2": 40, "D3": None}


def test_circular_reference_through_a_spill_is_refused(tmp_path):
    # B1 = FILTER(C1:C3, C1:C3>0) spills into B2, but C2 = B2 reads that spill: circular
    book = make_book(tmp_path / "p4.xlsx", {"C": [1, None, 2]},
                     {"B1": "=FILTER(C1:C3,C1:C3>0)", "C2": "=B2"})
    assert_refused(book, "Circular reference at L!B1")


def test_overlapping_spill_is_refused_for_the_later_formula(tmp_path):
    # B1 = SEQUENCE(3) claims B1:B3 first (row 1 is calculated first). A2 = SEQUENCE(2,2) would
    # need A2:B3 and B2 is already claimed, so A2 shows #SPILL! and writes nothing.
    book = make_book(tmp_path / "p5.xlsx", {},
                     {"B1": "=SEQUENCE(3)", "A2": "=SEQUENCE(2,2)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3", "A2", "A3"]) == {
        "B1": 1, "B2": 2, "B3": 3, "A2": "#SPILL!", "A3": None}


def test_array_comparison_and_elementwise_arithmetic(tmp_path):
    # A1:A3 * 2 is 2, 4, 6; A1:A3 > 1 is FALSE, TRUE, TRUE; keeping the TRUE rows gives 4, 6
    book = make_book(tmp_path / "p6.xlsx", {"A": [1, 2, 3]},
                     {"B1": "=FILTER(A1:A3*2,A1:A3>1)"})
    out = recalc_ok(book)
    assert read_back(out, ["B1", "B2", "B3"]) == {"B1": 4, "B2": 6, "B3": None}


# ---------------------------------------------------------------- write prefixes

def test_write_adds_the_excel_prefixes_for_the_new_functions(tmp_path):
    book = make_book(tmp_path / "w1.xlsx", {"A": [1, 2, 3]}, {})
    patch = tmp_path / "patch.json"
    patch.write_text(json.dumps([
        {"sheet": "L", "cell": "B1", "formula": "=SORT(A1:A3)"},
        {"sheet": "L", "cell": "B2", "formula": "=FILTER(A1:A3,A1:A3>1)"},
        {"sheet": "L", "cell": "B3", "formula": "=UNIQUE(A1:A3)"},
        {"sheet": "L", "cell": "B4", "formula": "=SEQUENCE(2)"},
        {"sheet": "L", "cell": "B5", "formula": "=LET(x,1,x+1)"},
    ]), encoding="utf-8")
    proc = run("write", book, "--patch", patch, "--json")
    assert proc.returncode == 0, proc.stderr
    out_book = tmp_path / "w1.write.xlsx"
    wb = openpyxl.load_workbook(out_book)
    assert [wb["L"][f"B{i}"].value for i in range(1, 6)] == [
        "=_xlfn._xlws.SORT(A1:A3)",
        "=_xlfn._xlws.FILTER(A1:A3,A1:A3>1)",
        "=_xlfn.UNIQUE(A1:A3)",
        "=_xlfn.SEQUENCE(2)",
        "=_xlfn.LET(_xlpm.x,1,_xlpm.x+1)",
    ]
    # the prefixed LET still calculates when the file is recalculated: x + 1 with x = 1 is 2
    out = recalc_ok(out_book)
    assert read_back(out, ["B5"]) == {"B5": 2}
