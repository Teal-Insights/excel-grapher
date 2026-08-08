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


def test_index_excel_range_row_zero_with_col_selects_column() -> None:
    base = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=3, end_col=3)
    result = index_excel_range(base, 0, 2)
    assert isinstance(result, ExcelRange)
    assert (result.start_row, result.start_col, result.end_row, result.end_col) == (1, 2, 3, 2)
