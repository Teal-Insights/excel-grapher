"""Array operators feed ndarray operands straight to the array paths.

An ndarray operand already is the buffer the fast path and the reference loops
want. Wrapping it in a `Grid` must not copy it into nested lists nor rebuild it
one `Grid.at` call at a time — that round trip dominated the wall clock of
large-array operator workloads.
"""

# ruff: noqa: E402
from __future__ import annotations

from typing import Any

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core.grid import Grid, Range
from excel_grapher.core.operator_thresholds import MIN_OPERATOR_FASTPATH_CELLS
from excel_grapher.core.operators import xl_concat, xl_eq, xl_mul
from excel_grapher.core.types import CellValue, XlError
from tests.unit.core.operators_test_helpers import array_tolist

LARGE = MIN_OPERATOR_FASTPATH_CELLS * 2


@pytest.fixture
def grid_at_calls(monkeypatch: pytest.MonkeyPatch) -> list[tuple[int, int]]:
    """Record every positional `Grid.at` access made during a test."""
    calls: list[tuple[int, int]] = []
    original = Grid.at

    def spy(self: Grid, row0: int, col0: int) -> Any:
        calls.append((row0, col0))
        return original(self, row0, col0)

    monkeypatch.setattr(Grid, "at", spy)
    return calls


def _numeric_column(size: int) -> Any:
    return np.array([[float(i)] for i in range(size)], dtype=object)


def test_ndarray_scalar_arithmetic_skips_per_cell_grid_access(
    grid_at_calls: list[tuple[int, int]],
) -> None:
    left = _numeric_column(LARGE)
    result = xl_mul(left, 2.0)
    assert grid_at_calls == []
    assert array_tolist(result)[3] == [6.0]


def test_ndarray_pair_compare_skips_per_cell_grid_access(
    grid_at_calls: list[tuple[int, int]],
) -> None:
    left = np.array([[str(i)] for i in range(LARGE)], dtype=object)
    right = _numeric_column(LARGE)
    result = xl_eq(left, right)
    assert grid_at_calls == []
    assert array_tolist(result) == [[False]] * LARGE


def test_fastpath_miss_on_ndarray_operands_skips_per_cell_grid_access(
    grid_at_calls: list[tuple[int, int]],
) -> None:
    """Mixed cell types miss every fast-path tier but still reuse the arrays."""
    left = np.array([["abc"]] + [[float(i)] for i in range(1, LARGE)], dtype=object)
    right = np.array([[0.0]] + [[float(i)] for i in range(1, LARGE)], dtype=object)
    result = xl_eq(left, right)
    assert grid_at_calls == []
    assert array_tolist(result)[0] == [False]
    assert array_tolist(result)[1] == [True]


def test_ndarray_concat_skips_per_cell_grid_access(
    grid_at_calls: list[tuple[int, int]],
) -> None:
    left = np.array([["a"]] * LARGE, dtype=object)
    right = np.array([["b"]] * LARGE, dtype=object)
    result = xl_concat(left, right)
    assert grid_at_calls == []
    assert array_tolist(result) == [["ab"]] * LARGE


def test_one_dimensional_ndarray_operand_is_treated_as_a_column() -> None:
    left = np.array([float(i) for i in range(LARGE)], dtype=object)
    assert array_tolist(xl_mul(left, 3.0))[2] == [6.0]


def test_typed_float_ndarray_operand_still_produces_object_cells() -> None:
    left = np.arange(LARGE, dtype=np.float64).reshape(LARGE, 1)
    rows = array_tolist(xl_mul(left, 2.0))
    assert rows[4] == [8.0]
    assert isinstance(rows[4][0], float)


def test_ndarray_operands_below_threshold_still_compare_elementwise() -> None:
    left = np.array([[1.0, 2.0]], dtype=object)
    right = np.array([[1.0, 3.0]], dtype=object)
    assert array_tolist(xl_eq(left, right)) == [[True, False]]


def test_range_operands_still_use_positional_grid_access(
    grid_at_calls: list[tuple[int, int]],
) -> None:
    def resolve(address: str) -> CellValue:
        return 2.0

    rng = Range("S", 1, 1, LARGE, 1, resolve)
    result = xl_mul(rng, 3.0)
    assert grid_at_calls != []
    assert array_tolist(result)[0] == [6.0]


def test_ndarray_shape_mismatch_still_returns_value() -> None:
    left = np.array([[1.0, 2.0]], dtype=object)
    right = np.array([[1.0, 2.0, 3.0]], dtype=object)
    assert xl_eq(left, right) == XlError.VALUE
