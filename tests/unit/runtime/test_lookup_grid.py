"""Grid-native INDEX and lookup wrappers for the evaluator runtime (#336 Phase 1)."""

from __future__ import annotations

from excel_grapher.core import CellValue, XlError
from excel_grapher.core.excel_function_meta import (
    eager_materialize_arg_indices,
    grid_range_arg_indices,
)
from excel_grapher.core.grid import Range
from excel_grapher.core.lookup_funcs import index_cells
from excel_grapher.runtime.lookup import xl_index


def test_index_not_in_eager_materialize_arg_indices() -> None:
    """INDEX uses geometry binding; it is not an eager ndarray consumer."""
    assert eager_materialize_arg_indices("INDEX") == frozenset()


def test_lookup_args_are_grid_range_bound() -> None:
    """Lookup table args bind lazy Range; other args do not."""
    assert 1 in grid_range_arg_indices("MATCH")
    assert 1 in grid_range_arg_indices("VLOOKUP")
    assert grid_range_arg_indices("ABS") == frozenset()


def test_index_cells_over_lazy_range_is_selective() -> None:
    calls: list[str] = []
    values: dict[str, CellValue] = {
        "S!A1": 10,
        "S!A2": 20,
        "S!A3": XlError.DIV,
    }

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return values[address]

    rng = Range("S", 1, 1, 3, 1, resolve)
    assert index_cells(rng, 2, None) == 20
    assert calls == ["S!A2"]


def test_xl_index_over_lazy_range_agrees_with_nested_list() -> None:
    values = {
        "S!A1": 1,
        "S!B1": 2,
        "S!A2": 3,
        "S!B2": 4,
    }

    def resolve(address: str) -> CellValue:
        return values[address]

    lazy = Range("S", 1, 1, 2, 2, resolve)
    eager = [[1, 2], [3, 4]]
    assert xl_index(lazy, 2, 1) == xl_index(eager, 2, 1) == 3
    assert xl_index(lazy, 1, 2) == 2
