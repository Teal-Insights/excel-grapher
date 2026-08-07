"""`Grid` keeps ndarray operands as arrays instead of eagerly copying nested rows."""

# ruff: noqa: E402
from __future__ import annotations

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core.grid import Grid, Range
from excel_grapher.core.types import CellValue


def test_wrap_ndarray_keeps_the_original_array() -> None:
    arr = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    grid = Grid.wrap(arr)
    assert grid is not None
    assert grid.array is arr
    assert (grid.nrows, grid.ncols) == (2, 2)


def test_wrap_ndarray_still_serves_positional_access() -> None:
    arr = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    grid = Grid.wrap(arr)
    assert grid is not None
    assert grid.at(1, 0) == 3.0
    assert grid.at_flat(1) == 2.0
    assert list(grid.iter_raw()) == [1.0, 2.0, 3.0, 4.0]


def test_wrap_one_dimensional_ndarray_is_a_column() -> None:
    arr = np.array([1.0, 2.0, 3.0], dtype=object)
    grid = Grid.wrap(arr)
    assert grid is not None
    assert (grid.nrows, grid.ncols) == (3, 1)
    assert grid.at(2, 0) == 3.0


def test_wrap_row_empty_ndarray_collapses_to_a_single_blank_cell() -> None:
    grid = Grid.wrap(np.empty((0, 3), dtype=object))
    assert grid is not None
    assert (grid.nrows, grid.ncols) == (1, 1)
    assert grid.at(0, 0) is None


def test_wrap_column_empty_ndarray_keeps_zero_width() -> None:
    grid = Grid.wrap(np.empty((2, 0), dtype=object))
    assert grid is not None
    assert (grid.nrows, grid.ncols) == (2, 0)


def test_ndarray_backed_row_and_column_slices_are_nested_lists() -> None:
    arr = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    grid = Grid.wrap(arr)
    assert grid is not None
    assert grid.row_slice(1) == [[3.0, 4.0]]
    assert grid.col_slice(0) == [[1.0], [3.0]]


def test_non_ndarray_grids_have_no_backing_array() -> None:
    def resolve(address: str) -> CellValue:
        return 1

    nested = Grid.wrap([[1, 2], [3, 4]])
    assert nested is not None
    assert nested.array is None

    ranged = Grid.wrap(Range("S", 1, 1, 2, 1, resolve))
    assert ranged is not None
    assert ranged.array is None


def test_wrap_still_returns_none_for_scalars() -> None:
    assert Grid.wrap(3) is None
    assert Grid.wrap("text") is None
    assert Grid.wrap(None) is None
