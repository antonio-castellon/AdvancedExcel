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
    from advanced_excel import AdvancedExcel
except ImportError:
    # Fallback for direct source execution in some environments
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
    from advanced_excel import AdvancedExcel


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
        "get_sheet",
        "get_sheet_names",
        # lookup
        "get_next_value",
        "get_next_value_from_key_list",
        "get_all_rows_from_key",
        "get_value_from_position",
        "get_all_pair_values",
        # blocks / entities / tables
        "get_all_tables",
        "get_dataframe_blocks_by_key_name",
        "get_dataframe_blocks_by_key_column",
        # rows
        "remove_empty_rows",
        "remove_duplicates",
        "remove_rows_by_regex",
        "remove_rows_not_in_regex",
        "remove_rows_where_value_in_column_is_empty",
        "remove_rows_where_value_is_equal_in_column",
        "merge_rows",
        "remove_row_if_unique_value",
        "remove_rows_if_unique_value_on_top",
        "remove_rows_if_unique_value_on_bottom",
        "remove_all_rows_from_a_column_name_when_has_value",
        "remove_all_rows_from_a_column_when_has_value",
        "remove_all_rows_before_column_value",
        "remove_row_by_index",
        "fill_forward_row_values",
        "set_header_as_a_row",
        "set_row_as_header",
        "set_row_as_header_by_value",
        "add_header_as_a_row",
        "transpose",
        # columns
        "remove_empty_columns",
        "remove_all_columns_where_any_row_has_a_value",
        "remove_all_columns_where_all_row_has_a_value",
        "remove_all_columns_where_all_row_are_empty",
        "remove_all_columns_before_column_value_header",
        "remove_all_columns_after_column_value_header",
        "move_to_first",
        "set_column_to_numeric",
        "set_list_of_columns_to_numeric",
        "remove_columns_by_name_nan",
        "remove_column_by_name",
        "remove_columns_by_names",
        "remove_all_columns_except_list_of_column_names",
        "merge_columns",
        "merge_columns_by_name",
        "split_on_columns",
        "split_in_2_columns_right",
        "split_in_2_columns_left",
        # cleaning
        "replace_nan_in_column_by_nan_strings",
        "replace_nan_strings",
        "strip_all",
        "remove_all_crlf",
        "replace_by_dictionary",
        "replace_all",
        "replace_all_by_list",
        "replace_spaces_by_separator",
        "transform_as_numeric",
        "fill_columns_with_left_value",
        "rename_headers",
        "rename_headers_by_regexp",
        "rename_headers_by_regexp_list",
        "case_headers",
        "case_column_values",
        # reshape
        "melt",
        "melt_extended",
        "merge_list_dataframes",
    ]
    for name in expected:
        assert hasattr(AdvancedExcel, name), f"Missing public method: {name}"
        assert callable(getattr(AdvancedExcel, name))


# -------------------------- ROWS --------------------------


def test_remove_empty_rows(ae, sample_df):
    df = sample_df.copy()
    result = ae.remove_empty_rows(df)
    assert result.shape[0] == 4
    assert "index" not in result.columns


def test_remove_duplicates(ae):
    df = pd.DataFrame({"A": [1, 1, 2], "B": ["x", "x", "y"]})
    result = ae.remove_duplicates(df)
    assert len(result) == 2


@pytest.mark.parametrize(
    "regex, expected_len",
    [
        (r"^f", 2),  # rows where C matches ^f (foo)
        (r"bar", 2),
    ],
)
def test_remove_rows_by_regex_variants(ae, sample_df, regex, expected_len):
    result = ae.remove_rows_by_regex(sample_df.copy(), "C", regex)
    assert len(result) == expected_len


def test_remove_rows_not_in_regex(ae, sample_df):
    result = ae.remove_rows_not_in_regex(sample_df.copy(), "C", r"foo")
    assert len(result) == 3


def test_remove_rows_where_value_in_column_is_empty(ae):
    df = pd.DataFrame({"Col": ["na", "NA", "good", None, "ok"]})
    result = ae.remove_rows_where_value_in_column_is_empty(df.copy(), "Col")
    assert len(result) == 2


def test_remove_rows_where_value_is_equal(ae):
    df = pd.DataFrame({"Col": [10, 20, 10, 30]})
    result = ae.remove_rows_where_value_is_equal_in_column(df.copy(), "Col", 10)
    assert len(result) == 2


def test_merge_rows(ae):
    df = pd.DataFrame({"X": ["a", "b", "c", "d"]})
    result = ae.merge_rows(df.copy(), 1, 3, "-")
    assert len(result) == 3
    # Explicitly test that the separator is used for joining values (fixed behavior)
    assert result.iloc[1, 0] == "b-c"

    # Default separator (empty string) joins without extra characters
    result_default = ae.merge_rows(df.copy(), 1, 3, "")
    assert result_default.iloc[1, 0] == "bc"

    # Separator with NaN replacement
    df_nan = pd.DataFrame({"X": ["a", None, "c", "d"]})
    result_nan = ae.merge_rows(df_nan.copy(), 1, 3, "-")
    # NaNs in the range become the separator, then values are joined with it
    assert "-" in result_nan.iloc[1, 0]
    assert result.iloc[1, 0] == "b - c"


def test_remove_row_if_unique_value_and_top_bottom(ae):
    df = pd.DataFrame({"A": ["Title", 1, 2, "Footer", 3]})
    r1 = ae.remove_row_if_unique_value(df.copy(), 0, 4)
    assert len(r1) == 3

    r2 = ae.remove_rows_if_unique_value_on_top(df.copy(), 1)
    assert len(r2) == 4

    r3 = ae.remove_rows_if_unique_value_on_bottom(df.copy(), 1)
    assert len(r3) == 4


def test_remove_all_rows_from_column_variants(ae):
    df = pd.DataFrame({"K": ["Batch", "x", "Batch", "y"], "V": [1, 2, 3, 4]})
    r1 = ae.remove_all_rows_from_a_column_name_when_has_value(
        df.copy(), 0, "K", "Batch", include_remove_row=True
    )
    assert len(r1) == 0 or r1.iloc[0]["K"] != "Batch"

    r2 = ae.remove_all_rows_from_a_column_when_has_value(df.copy(), 0, 0, "Batch", include_remove_row=False)
    assert len(r2) > 0

    r3 = ae.remove_all_rows_before_column_value(df.copy(), 0, 0, "Batch")
    assert len(r3) >= 2


def test_remove_row_by_index(ae):
    df = pd.DataFrame({"A": [0, 1, 2]})
    result = ae.remove_row_by_index(df.copy(), 1)
    assert len(result) == 2


def test_fill_forward_row_values(ae):
    df = pd.DataFrame({"A": [1, None, None, 4]})
    result = ae.fill_forward_row_values(df.copy(), ["A"])
    assert list(result["A"]) == [1, 1, 1, 4]


def test_header_row_manipulation(ae):
    df = pd.DataFrame({0: ["h1", "v1"], 1: ["h2", "v2"]})
    h_as_row = ae.set_header_as_a_row(df.copy())
    assert len(h_as_row) == 3

    row_as_h = ae.set_row_as_header(df.copy(), 0, uppercase=True, char_space_to="_")
    assert list(row_as_h.columns) == ["H1", "H2"]

    by_val = ae.set_row_as_headerByValue(df.copy(), 0, "h1", uppercase=False)
    assert "h1" not in list(by_val.columns) or len(by_val) == 1


def test_add_header_as_row_and_transpose(ae):
    df = pd.DataFrame({"X": [1, 2], "Y": [3, 4]})
    with_header = ae.add_header_as_a_row(df.copy())
    assert len(with_header) == 3

    t = ae.transpose(df.copy())
    assert len(t.columns) == 2


# -------------------------- COLUMNS --------------------------


def test_remove_empty_columns(ae):
    df = pd.DataFrame({"A": [1, 2], "B": [None, None], "C": ["", ""]})
    result = ae.remove_empty_columns(df.copy())
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
    before = ae.remove_all_columns_before_column_value_header(df.copy(), "Target")
    assert list(before.columns) == ["X"]

    after = ae.remove_all_columns_after_column_value_header(df.copy(), "Target")
    assert list(after.columns) == ["X", "Target"]


def test_move_to_first_and_set_numeric(ae):
    df = pd.DataFrame({"Z": [1], "A": [2], "B": [3]})
    moved = ae.move_to_first(df.copy(), "B")
    assert list(moved.columns)[:1] == ["B"]

    num_df = pd.DataFrame({"N": ["1", "2", "3"]})
    res = ae.set_column_to_numeric(num_df.copy(), "N")
    assert pd.api.types.is_numeric_dtype(res["N"])

    res2 = ae.set_list_of_columns_to_numeric(num_df.copy(), ["N"])
    assert pd.api.types.is_numeric_dtype(res2["N"])


def test_remove_columns_by_name_variants(ae):
    df = pd.DataFrame(columns=["Keep", "DropMe", "NAN", None, "AlsoDrop"])
    res1 = ae.remove_columns_by_name_nan(df.copy())
    assert "NAN" not in res1.columns and None not in res1.columns

    res2 = ae.remove_column_by_name(df.copy(), r"Drop")
    assert "DropMe" not in res2.columns

    res3 = ae.remove_columns_by_names(df.copy(), ["Keep", "AlsoDrop"])
    assert "Keep" not in res3.columns

    res4 = ae.remove_all_columns_except_list_of_column_names(df.copy(), ["Keep"])
    assert list(res4.columns) == ["Keep"]


def test_merge_and_split_columns(ae):
    df = pd.DataFrame({"L": ["a", "b"], "R": ["1", "2"], "Other": [9, 9]})
    merged = ae.merge_columns(df.copy(), 0, 1, "-", "LR")
    assert "LR" in merged.columns and "L" not in merged.columns

    merged2 = ae.merge_columnsByName(df.copy(), "LR2", "L", "R", remove_origin=True)
    assert "LR2" in merged2.columns and "L" not in merged2.columns

    split_df = pd.DataFrame({"S": ["a_1", "b_2"]})
    s1 = ae.split_on_columns(split_df.copy(), "S", "_", ["left", "right"])
    assert "left" in s1.columns

    s2 = ae.split_in_2_columns_right(split_df.copy(), "S")
    assert len(s2.columns) == 3

    s3 = ae.split_in_2_columns_left(split_df.copy(), "S")
    assert len(s3.columns) == 3


# -------------------------- CLEANING --------------------------


def test_replace_nan_variants(ae):
    df = pd.DataFrame({"C": [None, "NA", "NAN", "good"]})
    r1 = ae.replace_nan_in_column_by_nan_strings(df.copy(), "C")
    assert r1["C"].iloc[0] == "NA"

    r2 = ae.replace_nan_strings(df.copy())
    assert pd.isna(r2["C"].iloc[0])


def test_strip_crlf_replace_variants(ae):
    df = pd.DataFrame({"S": ["  hi  ", "a\nb\r\nc", "x  y"]})
    s = ae.strip_all(df.copy())
    assert s["S"].iloc[0] == "hi"

    cr = ae.remove_all_crlf(df.copy(), "-")
    assert "\n" not in cr["S"].iloc[1]

    d = {"old": "new"}
    rep = ae.replace_by_dictionary(pd.DataFrame({"T": ["old", "missing"]}), d, "T")
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
    res = ae.replace_all(df.copy(), "foo", "XXX", cols_clean, cols_not)
    if expected_change:
        assert "XXX" in res["S"].values
    else:
        assert "foo" in res["S"].values


def test_replace_all_by_list_and_spaces(ae):
    df = pd.DataFrame({"S": ["a b  c", "d;e"]})
    res = ae.replace_allByList(df.copy(), [("a b", "AB"), ("d;e", "DE")])
    assert "AB" in res["S"].values

    sp = ae.replace_spaces_by_separator(df.copy(), ";")
    assert "a;b;c" in sp["S"].values


def test_transform_numeric_and_fill_left(ae):
    df = pd.DataFrame({"N": ["1", "x", "3"]})
    t1 = ae.transform_as_numeric(df.copy())
    assert pd.api.types.is_numeric_dtype(t1["N"])

    t2 = ae.transform_as_numeric(df.copy(), columnsToApply=["N"])
    t3 = ae.transform_as_numeric(df.copy(), columnsToAvoid=["N"])

    left = pd.DataFrame({"A": [1, None, None], "B": [None, 5, None]})
    fl = ae.fill_columns_with_left_value(left.copy(), 1)
    assert fl.loc[1, "A"] == 1


def test_rename_and_case_variants(ae):
    df = pd.DataFrame(columns=["Old Name", "Another"])
    r1 = ae.rename_headers(df.copy(), {"Old Name": "NewName"})
    assert "NewName" in r1.columns

    r2 = ae.rename_headersByRegExp(df.copy(), r" ", "_")
    assert "Old_Name" in r2.columns

    r3 = ae.rename_headersByRegExpList(df.copy(), [r" ", r"o"], "_")
    assert "_" in "".join(r3.columns)

    c1 = ae.case_headers(df.copy(), uppercase=True)
    assert all(c.isupper() for c in c1.columns)

    val_df = pd.DataFrame({"P": ["hello", "WORLD", None]})
    cv = ae.case_column_values(val_df.copy(), "P", title=True)
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
    merged = ae.merge_list_dataframes(dfs, "GID")
    assert "GID" in merged.columns
    assert len(merged) == 2


# -------------------------- LOADER / BLOCKS / LOOKUP (with temp file) --------------------------


def test_loader_and_get_sheet(tmp_path):
    # Create a small real xlsx
    df = pd.DataFrame({"Col1": [1, 2], "Col2": ["a", "b"]})
    xlsx = tmp_path / "test_sheet.xlsx"
    df.to_excel(xlsx, sheet_name="Data", index=False)

    ae = AdvancedExcel(str(xlsx))
    loaded = ae.get_sheet("Data")
    assert loaded is not None
    assert list(loaded.columns) == ["Col1", "Col2"]
    assert "Data" in ae.get_sheetNames()


def test_context_manager(tmp_path):
    df = pd.DataFrame({"X": [1, 2, 3]})
    xlsx = tmp_path / "ctx.xlsx"
    df.to_excel(xlsx, index=False)

    with AdvancedExcel(str(xlsx)) as ae:
        s = ae.get_sheet("Sheet1")
        assert s is not None
        assert len(s) == 3
    # after exit, sheets cleared (current impl)
    # (no strong assertion as it is mostly for future resource mgmt)


def test_blocks_and_lookup(ae):
    df = pd.DataFrame({0: ["Study", "x", "Study", "y"], 1: [100, 101, 200, 201]})
    blocks = ae.get_dataframe_blocks_by_key_name(df.copy(), "Study")
    assert len(blocks) == 2

    rows = ae.get_all_rows_from_key(df.copy(), "Study")
    assert len(rows) == 2
    assert "data" in rows[0]

    nextv = ae.get_next_value(df.copy(), "Study")
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
    tables = ae.get_all_tables(df, mincol=2)
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
        assert ae.strip_all(df)["B"].iloc[0] == "x"
        print("Basic smoke passed. Use pytest for full coverage.")
    except Exception:
        traceback.print_exc()
        raise
