"""Unit tests for Grid-native flatten / aggregates (#336 Phase 3)."""

from __future__ import annotations

from typing import cast

import numpy as np

from excel_grapher.core import CellValue, XlError, flatten, get_error
from excel_grapher.core.excel_function_meta import (
    eager_materialize_arg_indices,
    grid_range_arg_indices,
)
from excel_grapher.core.grid import Range
from excel_grapher.core.math_funcs import countif_cells, sum_cells
from excel_grapher.core.sumproduct import sumproduct_cells


def test_full_scan_aggregates_bind_lazy_ranges_not_eager_ndarrays() -> None:
    """Phase 3 drops eager materialization for SUM / SUMPRODUCT / COUNTIF / AND / OR."""
    assert eager_materialize_arg_indices("SUM") == frozenset()
    assert eager_materialize_arg_indices("SUMPRODUCT") == frozenset()
    assert eager_materialize_arg_indices("COUNTIF") == frozenset()
    assert 0 in grid_range_arg_indices("SUM")
    assert 0 in grid_range_arg_indices("SUMPRODUCT")
    assert 0 in grid_range_arg_indices("COUNTIF")
    assert 0 in grid_range_arg_indices("AND")
    assert 0 in grid_range_arg_indices("OR")


def test_flatten_walks_lazy_range_in_row_major_order() -> None:
    calls: list[str] = []
    values = {"S!A1": 1, "S!B1": 2, "S!A2": 3, "S!B2": 4}

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return values[address]

    rng = Range("S", 1, 1, 2, 2, resolve)
    assert list(flatten(rng)) == [1, 2, 3, 4]
    assert calls == ["S!A1", "S!B1", "S!A2", "S!B2"]


def test_flatten_yields_error_sentinels_from_lazy_range() -> None:
    rng = Range(
        "S",
        1,
        1,
        3,
        1,
        lambda a: {"S!A1": 1, "S!A2": XlError.DIV, "S!A3": 3}[a],
    )
    assert list(flatten(rng)) == [1, XlError.DIV, 3]


def test_get_error_finds_first_error_in_lazy_range() -> None:
    """Full-scan precheck walks Range cells (lookups skip get_error instead)."""
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": 1, "S!A2": XlError.DIV, "S!A3": 3}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert get_error(rng) == XlError.DIV
    assert calls == ["S!A1", "S!A2"]


def test_sum_cells_over_lazy_range() -> None:
    rng = Range("S", 1, 1, 3, 1, lambda a: {"S!A1": 1, "S!A2": 2, "S!A3": 3}[a])
    assert sum_cells(cast(CellValue, rng)) == 6.0


def test_sum_cells_over_lazy_range_agrees_with_ndarray() -> None:
    values = {"S!A1": 1.5, "S!A2": 2.5, "S!A3": 3.5}

    def resolve(address: str) -> CellValue:
        return values[address]

    lazy = Range("S", 1, 1, 3, 1, resolve)
    eager = np.array([[1.5], [2.5], [3.5]], dtype=object)
    assert sum_cells(cast(CellValue, lazy)) == sum_cells(eager) == 7.5


def test_sum_cells_fail_fast_on_embedded_error() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": 1, "S!A2": XlError.DIV, "S!A3": 99}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert sum_cells(cast(CellValue, rng)) == XlError.DIV
    assert "S!A3" not in calls


def test_sumproduct_over_lazy_ranges() -> None:
    left = Range("S", 1, 1, 3, 1, lambda a: {"S!A1": 1, "S!A2": 2, "S!A3": 3}[a])
    right = Range("S", 1, 2, 3, 2, lambda a: {"S!B1": 4, "S!B2": 5, "S!B3": 6}[a])
    assert sumproduct_cells(cast(CellValue, left), cast(CellValue, right)) == 32.0


def test_sumproduct_shape_mismatch_returns_value_error() -> None:
    left = Range("S", 1, 1, 2, 1, lambda _a: 1)
    right = Range("S", 1, 1, 3, 1, lambda _a: 1)
    assert sumproduct_cells(cast(CellValue, left), cast(CellValue, right)) == XlError.VALUE


def test_countif_over_lazy_range() -> None:
    rng = Range(
        "S",
        1,
        1,
        3,
        1,
        lambda a: {"S!A1": 10, "S!A2": 3, "S!A3": 20}[a],
    )
    assert countif_cells(cast(CellValue, rng), ">5") == 2


def test_countif_skips_error_cells_in_lazy_range() -> None:
    """Excel COUNTIF ignores error cells; it does not propagate them."""
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": 10, "S!A2": XlError.DIV, "S!A3": 20}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert countif_cells(cast(CellValue, rng), ">5") == 2
    assert calls == ["S!A1", "S!A2", "S!A3"]
