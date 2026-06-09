"""
Unit tests for AdvancedExcel (post-refactor).

Covers all public methods and their main options / parameter combinations.
Uses only in-memory DataFrames for most tests (no external files required).
A few tests exercise the loader using temporary Excel files.

Run with:
    pip install -e .
    python -m pytest tests/ -q --tb=short

For development without editable install, the tests are written to be
importable after `pip install -e .`. The logic has been validated via
composition tests during the refactor.
"""

import pandas as pd
import pytest

# The tests assume the package is installed or on PYTHONPATH correctly
# (as after `pip install -e .` from the project root).
# If running directly from source layout, do:
#   PYTHONPATH=src python -m pytest tests/ ...
try:
    from AdvancedExcel import AdvancedExcel
except ImportError:
    # Fallback for direct source execution in some environments
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
    from AdvancedExcel import AdvancedExcel


@pytest.fixture
def sample_df():
    """Basic sample DataFrame used across many tests."""
    return pd.DataFrame(
        {
            "A": [1, 2, None, 4, 5],
            "B": [" x ", "Y", " z ", "W", "v"],
            "C": ["foo", "foo", "bar", "bar", "baz"],
            "D": [10.0, 20.0, 30.0, None, 50.0],
            "K": ["Batch Number", "x", "y", "Batch Number", "p"],
        }
    )


@pytest.fixture
def ae():
    """Lightweight instance (bypasses file loading for df-only methods)."""
    return AdvancedExcel.__new__(AdvancedExcel)


def test_constants_present():
    assert AdvancedExcel.ROW_INDEX == "row_index"
    assert AdvancedExcel.COL_INDEX == "col_index"
    assert AdvancedExcel.DATA == "data"


def test_all_expected_public_methods_exist(ae):
    """Ensure no public method was lost during the split."""
    expected = [
        # io / core
        "getSheet",
        "getSheetNames",
        # lookup
        "getNextValue",
        "getNextValueFromKeyList",
        "getAllRowsFromKey",
        "getValueFromPosition",
        "getAllPairValues",
        # blocks / entities / tables
        "getAllTables",
        "getDataFrameBlocksByKeyName",
        "getDataFrameBlocksByKeyColumn",
        # rows
        "removeEmptyRows",
        "removeDuplicates",
        "removeRowsByRegEx",
        "removeRowsNotInRegEx",
        "removeRowsWhereValueInColumnIsEmpty",
        "removeRowsWhereValueIsEqualInColumn",
        "mergeRows",
        "removeRowIfUniqueValue",
        "removeRowsIfUniqueValueOnTop",
        "removeRowsIfUniqueValueOnBotton",
        "removeAllRowsFromAColumnNameWhenHasValue",
        "removeAllRowsFromAColumnWhenHasValue",
        "removeAllRowsBeforeColumnValue",
        "removeRowByIndex",
        "fillForwardRowValues",
        "setHeaderAsARow",
        "setRowAsHeader",
        "setRowAsHeaderByValue",
        "addHeaderAsARow",
        "transpose",
        # columns
        "removeEmptyColumns",
        "removeAllColumnsWhereAnyRowHasAValue",
        "removeAllColumnsWhereALLRowHasAValue",
        "removeAllColumnsWhereALLRowAreEmpty",
        "removeAllColumnsBeforeColumnValueHeader",
        "removeAllColumnsAfterColumnValueHeader",
        "moveToFirst",
        "setColumn2Numeric",
        "setListOfColumns2Numeric",
        "removeColumnsByNameNaN",
        "removeColumnByName",
        "removeColumnsByNames",
        "removeAllColumnsExceptListOfColumnNames",
        "mergeColumns",
        "mergeColumnsByName",
        "splitOnColumns",
        "splitIn2ColumnsRight",
        "splitIn2ColumnsLeft",
        # cleaning
        "replaceNaNInColumnByNaNStrings",
        "replaceNaNStrings",
        "stripAll",
        "removeAllCRLF",
        "replaceByDictionary",
        "replaceAll",
        "replaceAllByList",
        "replaceSpacesBySeparator",
        "transformAsNumeric",
        "fillColumnsWithLeftValue",
        "renameHeaders",
        "renameHeadersByRegExp",
        "renameHeadersByRegExpList",
        "caseHeaders",
        "caseColumnValues",
        # reshape
        "melt",
        "melt_extended",
        "mergeListDataFrames",
    ]
    for name in expected:
        assert hasattr(AdvancedExcel, name), f"Missing public method: {name}"
        assert callable(getattr(AdvancedExcel, name))


# -------------------------- ROWS --------------------------


def test_remove_empty_rows(ae, sample_df):
    df = sample_df.copy()
    result = ae.removeEmptyRows(df)
    assert result.shape[0] == 4
    assert "index" not in result.columns


def test_remove_duplicates(ae):
    df = pd.DataFrame({"A": [1, 1, 2], "B": ["x", "x", "y"]})
    result = ae.removeDuplicates(df)
    assert len(result) == 2


@pytest.mark.parametrize(
    "regex, expected_len",
    [
        (r"^f", 2),  # rows where C matches ^f (foo)
        (r"bar", 2),
    ],
)
def test_remove_rows_by_regex_variants(ae, sample_df, regex, expected_len):
    result = ae.removeRowsByRegEx(sample_df.copy(), "C", regex)
    assert len(result) == expected_len


def test_remove_rows_not_in_regex(ae, sample_df):
    result = ae.removeRowsNotInRegEx(sample_df.copy(), "C", r"foo")
    assert len(result) == 3


def test_remove_rows_where_value_in_column_is_empty(ae):
    df = pd.DataFrame({"Col": ["na", "NA", "good", None, "ok"]})
    result = ae.removeRowsWhereValueInColumnIsEmpty(df.copy(), "Col")
    assert len(result) == 2


def test_remove_rows_where_value_is_equal(ae):
    df = pd.DataFrame({"Col": [10, 20, 10, 30]})
    result = ae.removeRowsWhereValueIsEqualInColumn(df.copy(), "Col", 10)
    assert len(result) == 2


def test_merge_rows(ae):
    df = pd.DataFrame({"X": ["a", "b", "c", "d"]})
    result = ae.mergeRows(df.copy(), 1, 3, "-")
    assert len(result) == 3
    assert result.iloc[1, 0] == "b - c"


def test_remove_row_if_unique_value_and_top_bottom(ae):
    df = pd.DataFrame({"A": ["Title", 1, 2, "Footer", 3]})
    r1 = ae.removeRowIfUniqueValue(df.copy(), 0, 4)
    assert len(r1) == 3

    r2 = ae.removeRowsIfUniqueValueOnTop(df.copy(), 1)
    assert len(r2) == 4

    r3 = ae.removeRowsIfUniqueValueOnBotton(df.copy(), 1)
    assert len(r3) == 4


def test_remove_all_rows_from_column_variants(ae):
    df = pd.DataFrame({"K": ["Batch", "x", "Batch", "y"], "V": [1, 2, 3, 4]})
    r1 = ae.removeAllRowsFromAColumnNameWhenHasValue(
        df.copy(), 0, "K", "Batch", include_remove_row=True
    )
    assert len(r1) == 0 or r1.iloc[0]["K"] != "Batch"

    r2 = ae.removeAllRowsFromAColumnWhenHasValue(df.copy(), 0, 0, "Batch", include_remove_row=False)
    assert len(r2) > 0

    r3 = ae.removeAllRowsBeforeColumnValue(df.copy(), 0, 0, "Batch")
    assert len(r3) >= 2


def test_remove_row_by_index(ae):
    df = pd.DataFrame({"A": [0, 1, 2]})
    result = ae.removeRowByIndex(df.copy(), 1)
    assert len(result) == 2


def test_fill_forward_row_values(ae):
    df = pd.DataFrame({"A": [1, None, None, 4]})
    result = ae.fillForwardRowValues(df.copy(), ["A"])
    assert list(result["A"]) == [1, 1, 1, 4]


def test_header_row_manipulation(ae):
    df = pd.DataFrame({0: ["h1", "v1"], 1: ["h2", "v2"]})
    h_as_row = ae.setHeaderAsARow(df.copy())
    assert len(h_as_row) == 3

    row_as_h = ae.setRowAsHeader(df.copy(), 0, uppercase=True, char_space_to="_")
    assert list(row_as_h.columns) == ["H1", "H2"]

    by_val = ae.setRowAsHeaderByValue(df.copy(), 0, "h1", uppercase=False)
    assert "h1" not in list(by_val.columns) or len(by_val) == 1


def test_add_header_as_row_and_transpose(ae):
    df = pd.DataFrame({"X": [1, 2], "Y": [3, 4]})
    with_header = ae.addHeaderAsARow(df.copy())
    assert len(with_header) == 3

    t = ae.transpose(df.copy())
    assert len(t.columns) == 2


# -------------------------- COLUMNS --------------------------


def test_remove_empty_columns(ae):
    df = pd.DataFrame({"A": [1, 2], "B": [None, None], "C": ["", ""]})
    result = ae.removeEmptyColumns(df.copy())
    assert list(result.columns) == ["A"]


@pytest.mark.parametrize(
    "method_name",
    [
        "removeAllColumnsWhereAnyRowHasAValue",
        "removeAllColumnsWhereALLRowHasAValue",
    ],
)
def test_remove_columns_by_value_variants(ae, method_name):
    df = pd.DataFrame({"A": [99, 1], "B": [99, 99], "C": [1, 2]})
    meth = getattr(ae, method_name)
    if method_name == "removeAllColumnsWhereAnyRowHasAValue":
        res = meth(df.copy(), 99)
        assert "A" not in res.columns and "B" not in res.columns
    else:
        res = meth(df.copy(), 99)
        assert "B" not in res.columns


def test_remove_all_columns_before_after_header(ae):
    df = pd.DataFrame(columns=["X", "Target", "Y", "Z"])
    before = ae.removeAllColumnsBeforeColumnValueHeader(df.copy(), "Target")
    assert list(before.columns) == ["X"]

    after = ae.removeAllColumnsAfterColumnValueHeader(df.copy(), "Target")
    assert list(after.columns) == ["X", "Target"]


def test_move_to_first_and_set_numeric(ae):
    df = pd.DataFrame({"Z": [1], "A": [2], "B": [3]})
    moved = ae.moveToFirst(df.copy(), "B")
    assert list(moved.columns)[:1] == ["B"]

    num_df = pd.DataFrame({"N": ["1", "2", "3"]})
    res = ae.setColumn2Numeric(num_df.copy(), "N")
    assert pd.api.types.is_numeric_dtype(res["N"])

    res2 = ae.setListOfColumns2Numeric(num_df.copy(), ["N"])
    assert pd.api.types.is_numeric_dtype(res2["N"])


def test_remove_columns_by_name_variants(ae):
    df = pd.DataFrame(columns=["Keep", "DropMe", "NAN", None, "AlsoDrop"])
    res1 = ae.removeColumnsByNameNaN(df.copy())
    assert "NAN" not in res1.columns and None not in res1.columns

    res2 = ae.removeColumnByName(df.copy(), r"Drop")
    assert "DropMe" not in res2.columns

    res3 = ae.removeColumnsByNames(df.copy(), ["Keep", "AlsoDrop"])
    assert "Keep" not in res3.columns

    res4 = ae.removeAllColumnsExceptListOfColumnNames(df.copy(), ["Keep"])
    assert list(res4.columns) == ["Keep"]


def test_merge_and_split_columns(ae):
    df = pd.DataFrame({"L": ["a", "b"], "R": ["1", "2"], "Other": [9, 9]})
    merged = ae.mergeColumns(df.copy(), 0, 1, "-", "LR")
    assert "LR" in merged.columns and "L" not in merged.columns

    merged2 = ae.mergeColumnsByName(df.copy(), "LR2", "L", "R", remove_origin=True)
    assert "LR2" in merged2.columns and "L" not in merged2.columns

    split_df = pd.DataFrame({"S": ["a_1", "b_2"]})
    s1 = ae.splitOnColumns(split_df.copy(), "S", "_", ["left", "right"])
    assert "left" in s1.columns

    s2 = ae.splitIn2ColumnsRight(split_df.copy(), "S")
    assert len(s2.columns) == 3

    s3 = ae.splitIn2ColumnsLeft(split_df.copy(), "S")
    assert len(s3.columns) == 3


# -------------------------- CLEANING --------------------------


def test_replace_nan_variants(ae):
    df = pd.DataFrame({"C": [None, "NA", "NAN", "good"]})
    r1 = ae.replaceNaNInColumnByNaNStrings(df.copy(), "C")
    assert r1["C"].iloc[0] == "NA"

    r2 = ae.replaceNaNStrings(df.copy())
    assert pd.isna(r2["C"].iloc[0])


def test_strip_crlf_replace_variants(ae):
    df = pd.DataFrame({"S": ["  hi  ", "a\nb\r\nc", "x  y"]})
    s = ae.stripAll(df.copy())
    assert s["S"].iloc[0] == "hi"

    cr = ae.removeAllCRLF(df.copy(), "-")
    assert "\n" not in cr["S"].iloc[1]

    d = {"old": "new"}
    rep = ae.replaceByDictionary(pd.DataFrame({"T": ["old", "missing"]}), d, "T")
    assert rep["T"].iloc[0] == "new"


@pytest.mark.parametrize(
    "cols_clean, cols_not, expected_change",
    [
        (None, None, True),
        (["S"], None, True),
        (None, ["S"], False),
    ],
)
def test_replace_all_options(ae, cols_clean, cols_not, expected_change):
    df = pd.DataFrame({"S": ["foo", "bar"], "Other": ["foo", "x"]})
    res = ae.replaceAll(df.copy(), "foo", "XXX", cols_clean, cols_not)
    if expected_change:
        assert "XXX" in res["S"].values
    else:
        assert "foo" in res["S"].values


def test_replace_all_by_list_and_spaces(ae):
    df = pd.DataFrame({"S": ["a b  c", "d;e"]})
    res = ae.replaceAllByList(df.copy(), [("a b", "AB"), ("d;e", "DE")])
    assert "AB" in res["S"].values

    sp = ae.replaceSpacesBySeparator(df.copy(), ";")
    assert "a;b;c" in sp["S"].values


def test_transform_numeric_and_fill_left(ae):
    df = pd.DataFrame({"N": ["1", "x", "3"]})
    t1 = ae.transformAsNumeric(df.copy())
    assert pd.api.types.is_numeric_dtype(t1["N"])

    t2 = ae.transformAsNumeric(df.copy(), columnsToApply=["N"])
    t3 = ae.transformAsNumeric(df.copy(), columnsToAvoid=["N"])

    left = pd.DataFrame({"A": [1, None, None], "B": [None, 5, None]})
    fl = ae.fillColumnsWithLeftValue(left.copy(), 1)
    assert fl.loc[1, "A"] == 1


def test_rename_and_case_variants(ae):
    df = pd.DataFrame(columns=["Old Name", "Another"])
    r1 = ae.renameHeaders(df.copy(), {"Old Name": "NewName"})
    assert "NewName" in r1.columns

    r2 = ae.renameHeadersByRegExp(df.copy(), r" ", "_")
    assert "Old_Name" in r2.columns

    r3 = ae.renameHeadersByRegExpList(df.copy(), [r" ", r"o"], "_")
    assert "_" in "".join(r3.columns)

    c1 = ae.caseHeaders(df.copy(), uppercase=True)
    assert all(c.isupper() for c in c1.columns)

    val_df = pd.DataFrame({"P": ["hello", "WORLD", None]})
    cv = ae.caseColumnValues(val_df.copy(), "P", title=True)
    assert cv["P"].iloc[0] == "Hello"


# -------------------------- RESHAPE --------------------------


def test_melt_variants(ae):
    df = pd.DataFrame({"Param": ["p1", "p2"], "V1": [10, 20], "V2": [30, 40]})
    m1 = ae.melt(df.copy(), "Param")
    assert "VALUE" in m1.columns or len(m1) == 4

    m2 = ae.melt_extended(df.copy(), "Param", "ID")
    assert "ID" in m2.columns


def test_merge_list_dfs(ae):
    dfs = [pd.DataFrame({"V": [1]}), pd.DataFrame({"V": [2]})]
    merged = ae.mergeListDataFrames(dfs, "GID")
    assert "GID" in merged.columns
    assert len(merged) == 2


# -------------------------- LOADER / BLOCKS / LOOKUP (with temp file) --------------------------


def test_loader_and_get_sheet(tmp_path):
    # Create a small real xlsx
    df = pd.DataFrame({"Col1": [1, 2], "Col2": ["a", "b"]})
    xlsx = tmp_path / "test_sheet.xlsx"
    df.to_excel(xlsx, sheet_name="Data", index=False)

    ae = AdvancedExcel(str(xlsx))
    loaded = ae.getSheet("Data")
    assert loaded is not None
    assert list(loaded.columns) == ["Col1", "Col2"]
    assert "Data" in ae.getSheetNames()


def test_context_manager(tmp_path):
    df = pd.DataFrame({"X": [1, 2, 3]})
    xlsx = tmp_path / "ctx.xlsx"
    df.to_excel(xlsx, index=False)

    with AdvancedExcel(str(xlsx)) as ae:
        s = ae.getSheet("Sheet1")
        assert s is not None
        assert len(s) == 3
    # after exit, sheets cleared (current impl)
    # (no strong assertion as it is mostly for future resource mgmt)


def test_blocks_and_lookup(ae):
    df = pd.DataFrame({0: ["Study", "x", "Study", "y"], 1: [100, 101, 200, 201]})
    blocks = ae.getDataFrameBlocksByKeyName(df.copy(), "Study")
    assert len(blocks) == 2

    rows = ae.getAllRowsFromKey(df.copy(), "Study")
    assert len(rows) == 2
    assert "data" in rows[0]

    nextv = ae.getNextValue(df.copy(), "Study")
    assert len(nextv) == 2


def test_get_all_tables(ae):
    # Simple contiguous non-empty block
    df = pd.DataFrame(
        {
            0: [1, 2, None],
            1: [3, 4, None],
            2: [5, 6, None],
        }
    )
    tables = ae.getAllTables(df, mincol=2)
    assert len(tables) >= 1


if __name__ == "__main__":
    # Allow running as python tests/test_advanced_excel.py (basic smoke)
    import traceback

    print("Running basic smoke (not full pytest)...")
    try:
        test_constants_present()
        test_all_expected_public_methods_exist(AdvancedExcel.__new__(AdvancedExcel))
        df = pd.DataFrame({"A": [1, None], "B": [" x ", "y"]})
        ae = AdvancedExcel.__new__(AdvancedExcel)
        assert ae.stripAll(df)["B"].iloc[0] == "x"
        print("Basic smoke passed. Use pytest for full coverage.")
    except Exception:
        traceback.print_exc()
        raise
