"""INDEX row_num/col_num 0 selects whole column/row (issues #502 / #503)."""

from __future__ import annotations

from excel_grapher.core.addressing import index_excel_range
from excel_grapher.core.grid import Range
from excel_grapher.core.lookup_funcs import index_cells, match_cells
from excel_grapher.core.types import CellValue, ExcelRange, XlError


def test_index_cells_row_zero_returns_whole_column() -> None:
    assert index_cells([[5], [0], [7]], 0) == [[5], [0], [7]]
    assert index_cells([[5, 6], [0, 1], [7, 8]], 0, 2) == [[6], [1], [8]]


def test_index_cells_col_zero_returns_whole_row() -> None:
    assert index_cells([[5, 6, 7]], None, 0) == [[5, 6, 7]]
    assert index_cells([[5, 6], [0, 1], [7, 8]], 2, 0) == [[0, 1]]


def test_index_cells_both_zero_returns_whole_array() -> None:
    array = [[5, 6], [0, 1], [7, 8]]
    assert index_cells(array, 0, 0) == array


def test_index_cells_row_zero_over_lazy_range() -> None:
    values = {"S!A1": 5, "S!A2": 0, "S!A3": 7}

    def resolve(address: str) -> CellValue:
        return values[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    result = index_cells(rng, 0)
    assert isinstance(result, Range)
    assert (result.start_row, result.end_row, result.start_col, result.end_col) == (1, 3, 1, 1)


def test_index_cells_negative_still_ref() -> None:
    assert index_cells([[5], [0], [7]], -1) == XlError.REF
    assert index_cells([[5, 6], [0, 1]], 1, -1) == XlError.REF


def test_match_true_over_index_boolean_column() -> None:
    """MATCH(TRUE, INDEX((rng<>0), 0), 0) idiom after materializing the compare."""
    booleans = [[True], [False], [True]]
    assert match_cells(True, index_cells(booleans, 0), 0) == 1


def test_index_excel_range_row_zero_returns_full_column() -> None:
    base = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=3, end_col=1)
    result = index_excel_range(base, 0, None)
    assert isinstance(result, ExcelRange)
    assert (result.start_row, result.start_col, result.end_row, result.end_col) == (1, 1, 3, 1)


def test_index_excel_range_col_zero_returns_full_row() -> None:
    base = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=3, end_col=3)
    result = index_excel_range(base, 2, 0)
    assert isinstance(result, ExcelRange)
    assert (result.start_row, result.start_col, result.end_row, result.end_col) == (2, 1, 2, 3)


def test_index_one_row_empty_column_is_the_row_and_two_arg_block_is_ref() -> None:
    """Excel 16: INDEX(A1:T1,1,) is the row; INDEX(A1:T3,1) is #REF!.

    `col_num is None` is the two-argument form. Callers pass `0` for an empty
    third argument, which matches `INDEX(...,1,0)`.
    """
    header = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=1, end_col=20)
    two_arg_header = index_excel_range(header, 1, None)
    assert isinstance(two_arg_header, ExcelRange)
    assert (two_arg_header.start_col, two_arg_header.end_col) == (1, 1)
    empty_header = index_excel_range(header, 1, 0)
    assert isinstance(empty_header, ExcelRange)
    assert (empty_header.start_col, empty_header.end_col) == (1, 20)

    block = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=3, end_col=20)
    assert index_excel_range(block, 1, None) == XlError.REF
    empty_row = index_excel_range(block, 1, 0)
    assert isinstance(empty_row, ExcelRange)
    assert (empty_row.start_row, empty_row.end_row, empty_row.start_col, empty_row.end_col) == (
        1,
        1,
        1,
        20,
    )

    assert index_cells([[1, 2, 3]], 1, None) == 1
    assert index_cells([[1, 2, 3]], 1, 0) == [[1, 2, 3]]
    assert index_cells([[1, 2], [3, 4]], 1, None) == XlError.REF
    assert index_cells([[1, 2], [3, 4]], 1, 0) == [[1, 2]]


def test_index_excel_range_row_zero_with_col_selects_column() -> None:
    base = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=3, end_col=3)
    result = index_excel_range(base, 0, 2)
    assert isinstance(result, ExcelRange)
    assert (result.start_row, result.start_col, result.end_row, result.end_col) == (1, 2, 3, 2)
