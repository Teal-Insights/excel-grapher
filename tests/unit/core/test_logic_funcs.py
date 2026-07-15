"""Unit tests for cell-wise AND/OR over lazy Range (#397)."""

# ruff: noqa: E402
from __future__ import annotations

from typing import cast

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.excel_function_meta import grid_range_arg_indices
from excel_grapher.core.grid import Range
from excel_grapher.core.logic_funcs import logical_and, logical_or


def test_and_or_bind_lazy_ranges() -> None:
    assert 0 in grid_range_arg_indices("AND")
    assert 0 in grid_range_arg_indices("OR")


def test_logical_and_over_lazy_range() -> None:
    rng = Range("S", 1, 1, 3, 1, lambda a: {"S!A1": True, "S!A2": True, "S!A3": True}[a])
    assert logical_and(cast(CellValue, rng)) is True


def test_logical_and_short_circuits_on_false_without_trailing_cells() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": True, "S!A2": False, "S!A3": True}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert logical_and(cast(CellValue, rng)) is False
    assert calls == ["S!A1", "S!A2"]


def test_logical_or_short_circuits_on_true_without_trailing_cells() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": False, "S!A2": True, "S!A3": False}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert logical_or(cast(CellValue, rng)) is True
    assert calls == ["S!A1", "S!A2"]


def test_logical_and_propagates_first_error_in_range() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": True, "S!A2": XlError.DIV, "S!A3": False}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert logical_and(cast(CellValue, rng)) == XlError.DIV
    assert calls == ["S!A1", "S!A2"]


def test_logical_or_propagates_first_error_in_range() -> None:
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return {"S!A1": False, "S!A2": XlError.DIV, "S!A3": True}[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert logical_or(cast(CellValue, rng)) == XlError.DIV
    assert calls == ["S!A1", "S!A2"]


def test_logical_and_skips_blank_cells() -> None:
    rng = Range(
        "S",
        1,
        1,
        3,
        1,
        lambda a: {"S!A1": None, "S!A2": True, "S!A3": None}[a],
    )
    assert logical_and(cast(CellValue, rng)) is True


def test_logical_or_skips_blank_cells() -> None:
    rng = Range(
        "S",
        1,
        1,
        3,
        1,
        lambda a: {"S!A1": None, "S!A2": False, "S!A3": None}[a],
    )
    assert logical_or(cast(CellValue, rng)) is False


def test_logical_and_all_blanks_returns_value_error() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda _a: None)
    assert logical_and(cast(CellValue, rng)) == XlError.VALUE


def test_logical_or_all_blanks_returns_value_error() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda _a: None)
    assert logical_or(cast(CellValue, rng)) == XlError.VALUE


def test_logical_and_invalid_text_returns_value_error() -> None:
    rng = Range("S", 1, 1, 1, 1, lambda _a: "not-a-bool")
    assert logical_and(cast(CellValue, rng)) == XlError.VALUE


def test_logical_and_coerces_numbers() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": 1, "S!A2": 0}[a])
    assert logical_and(cast(CellValue, rng)) is False


def test_logical_or_coerces_numbers() -> None:
    rng = Range("S", 1, 1, 2, 1, lambda a: {"S!A1": 0, "S!A2": 5}[a])
    assert logical_or(cast(CellValue, rng)) is True


def test_logical_and_scalar_args_unchanged() -> None:
    assert logical_and(True, True) is True
    assert logical_and(True, False) is False
    assert logical_or(True, False) is True
    assert logical_or(False, False) is False


def test_logical_and_over_ndarray_matches_range() -> None:
    values = {"S!A1": 1, "S!A2": 0, "S!A3": 2}

    def resolve(address: str) -> CellValue:
        return values[address]

    lazy = Range("S", 1, 1, 3, 1, resolve)
    eager = np.array([[1], [0], [2]], dtype=object)
    assert logical_and(cast(CellValue, lazy)) == logical_and(eager) is False
