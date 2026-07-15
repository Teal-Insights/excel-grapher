"""Unit tests for shared Grid operator maps (#336 Phase 2)."""

from __future__ import annotations

from typing import Any, cast

import pytest

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.grid import Range
from excel_grapher.core.operator_maps import (
    map_arithmetic,
    map_compare,
    map_concat,
    map_unary,
)
from excel_grapher.core.operators import xl_mul


def test_map_arithmetic_over_lazy_range_visits_each_cell_once() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": 1, "S!A2": 2, "S!A3": 3}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert map_arithmetic("*", rng, 2) == [[2], [4], [6]]
    assert calls == ["S!A1", "S!A2", "S!A3"]


def test_map_arithmetic_agrees_with_ndarray_xl_mul_when_fully_consumed() -> None:
    values = {"S!A1": 1.5, "S!A2": 2.5, "S!B1": 3.0, "S!B2": 4.0}

    def resolve(address: str) -> CellValue:
        return values[address]

    left = Range("S", 1, 1, 2, 1, resolve)
    right = Range("S", 1, 2, 2, 2, resolve)
    mapped = map_arithmetic("*", left, right)
    import numpy as np

    arr = xl_mul(
        np.array([[1.5], [2.5]], dtype=object),
        np.array([[3.0], [4.0]], dtype=object),
    )
    assert isinstance(arr, (np.ndarray, list))
    assert mapped == cast(Any, arr).tolist() if hasattr(arr, "tolist") else arr


def test_map_arithmetic_shape_mismatch_returns_value_error() -> None:
    left = Range("S", 1, 1, 2, 1, lambda _a: 1)
    right = Range("S", 1, 1, 3, 1, lambda _a: 1)
    assert map_arithmetic("+", left, right) == XlError.VALUE


def test_map_arithmetic_fail_fast_on_embedded_cell_error() -> None:
    def resolve(address: str) -> CellValue:
        return {"S!A1": 1, "S!A2": XlError.DIV, "S!A3": 3}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert map_arithmetic("*", rng, 2) == XlError.DIV


def test_map_compare_and_concat_over_nested_lists() -> None:
    assert map_compare("=", [["a"], ["b"]], "b") == [[False], [True]]
    assert map_concat([["x"], ["y"]], "!") == [["x!"], ["y!"]]


def test_map_unary_over_lazy_range() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": 4, "S!A2": -6}[a])
    assert map_unary("-", rng) == [[-4], [6]]
    assert map_unary("%", [[50], [25]]) == [[0.5], [0.25]]


def test_large_range_fastpath_miss_materializes_cells_only_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """When the vectorized fast path declines, reuse the materialized arrays."""
    from excel_grapher.core import operators as operators_mod
    from excel_grapher.core.operators_fastpath import MIN_OPERATOR_FASTPATH_CELLS

    nrows = MIN_OPERATOR_FASTPATH_CELLS
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        row = int(address.split("!")[1][1:])
        return float(row)

    monkeypatch.setattr(
        operators_mod,
        "try_fastpath_arithmetic_array",
        lambda *_args, **_kwargs: None,
    )

    rng = Range("S", 1, 1, nrows, 1, resolve)
    result = xl_mul(cast(CellValue, rng), 2.0)
    expected = [[float(i) * 2.0] for i in range(1, nrows + 1)]
    actual = cast(Any, result).tolist() if hasattr(result, "tolist") else result
    assert actual == expected
    assert len(calls) == nrows
    assert calls == [f"S!A{i}" for i in range(1, nrows + 1)]
