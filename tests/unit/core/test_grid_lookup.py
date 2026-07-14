"""Unit tests for shared lazy Range/Grid and lookup_funcs (#336)."""

from __future__ import annotations

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.grid import Grid, Range
from excel_grapher.core.lookup_funcs import (
    hlookup_cells,
    lookup_cells,
    match_cells,
    vlookup_cells,
    xlookup_cells,
)
from excel_grapher.core.types import XlErrorException


def test_core_range_value_at_keeps_error_sentinels() -> None:
    def resolve(address: str) -> CellValue:
        return {"S!A1": 1, "S!A2": XlError.DIV}[address]

    rng = Range("S", 1, 1, 2, 1, resolve)
    assert rng.value_at(1, 1) == 1
    assert rng.value_at(2, 1) == XlError.DIV


def test_core_range_cell_raises_on_error_sentinel() -> None:
    def resolve(address: str) -> CellValue:
        return XlError.NA

    rng = Range("S", 1, 1, 1, 1, resolve)
    try:
        rng.cell(1, 1)
    except XlErrorException as exc:
        assert exc.code == XlError.NA
    else:
        raise AssertionError("expected XlErrorException")


def test_grid_wrap_range_and_nested_lists() -> None:
    def resolve(address: str) -> CellValue:
        return {"S!A1": 1, "S!B1": 2}[address]

    rng = Range("S", 1, 1, 1, 2, resolve)
    grid = Grid.wrap(rng)
    assert grid is not None
    assert grid.at(0, 1) == 2

    list_grid = Grid.wrap([[10, 20], [30, 40]])
    assert list_grid is not None
    assert list_grid.at_flat(2) == 30
    assert Grid.wrap(5) is None


def test_match_vlookup_lazy_range_agrees_with_materialized_list() -> None:
    """Lazy Range and fully materialized nested lists agree when all cells are read."""
    values = {
        "S!A1": "k1",
        "S!A2": "k2",
        "S!A3": "k3",
        "S!B1": 10,
        "S!B2": 20,
        "S!B3": 30,
    }
    calls: list[str] = []

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return values[address]

    lazy = Range("S", 1, 1, 3, 2, resolve)
    eager = [
        [values["S!A1"], values["S!B1"]],
        [values["S!A2"], values["S!B2"]],
        [values["S!A3"], values["S!B3"]],
    ]

    assert match_cells("k2", lazy.column(1), 0) == match_cells("k2", [[row[0]] for row in eager], 0)
    assert vlookup_cells("k3", lazy, 2, False) == vlookup_cells("k3", eager, 2, False)
    assert set(calls) == {"S!A1", "S!A2", "S!A3", "S!B3"}


def test_hlookup_lookup_xlookup_lazy_agrees_with_materialized_list() -> None:
    values = {
        "S!A1": "k1",
        "S!B1": "k2",
        "S!C1": "k3",
        "S!A2": 1,
        "S!B2": 2,
        "S!C2": 3,
    }

    def resolve(address: str) -> CellValue:
        return values[address]

    lazy = Range("S", 1, 1, 2, 3, resolve)
    eager = [
        [values["S!A1"], values["S!B1"], values["S!C1"]],
        [values["S!A2"], values["S!B2"], values["S!C2"]],
    ]
    keys = lazy.row(1)
    vals = lazy.row(2)
    eager_keys = [eager[0]]
    eager_vals = [eager[1]]

    assert hlookup_cells("k2", lazy, 2, False) == hlookup_cells("k2", eager, 2, False)
    assert lookup_cells(20, [[10], [20], [30]], [["a"], ["b"], ["c"]]) == "b"
    assert xlookup_cells("k3", keys, vals) == xlookup_cells("k3", eager_keys, eager_vals)


def test_exact_match_on_lazy_range_stops_before_trailing_cells() -> None:
    calls: list[str] = []
    values: dict[str, CellValue] = {
        "S!A1": "x",
        "S!A2": "y",
        "S!A3": XlError.DIV,
    }

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return values[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert match_cells("y", rng, 0) == 2
    assert calls == ["S!A1", "S!A2"]
