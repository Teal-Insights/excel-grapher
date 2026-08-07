"""Scalar boundary and shallow get_error for lazy Range (#336 Phase 1)."""

# ruff: noqa: E402
from __future__ import annotations

from typing import cast

import pytest

np = pytest.importorskip("numpy")

from excel_grapher import DependencyGraph, Node
from excel_grapher.core import CellValue, XlError, as_scalar, get_error, to_number, to_string
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.grid import Range
from excel_grapher.core.lookup_funcs import index_cells
from excel_grapher.core.types import XlErrorException
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.export_runtime.lookup import xl_index as export_xl_index


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_as_scalar_rejects_range_list_and_ndarray() -> None:
    def resolve(address: str) -> CellValue:
        return 1

    rng = Range("S", 1, 1, 2, 1, resolve)
    assert as_scalar(rng) == XlError.VALUE
    assert as_scalar([[1], [2]]) == XlError.VALUE
    assert as_scalar(np.array([[1], [2]], dtype=object)) == XlError.VALUE
    assert as_scalar(3) == 3
    assert as_scalar(None) is None


def test_as_scalar_passes_plain_scalars_through_unchanged() -> None:
    """The plain-scalar shortcut must not alter the value or its type."""
    for value in (3, 3.5, "text", True, False, None):
        assert as_scalar(value) is value


def test_as_scalar_still_classifies_scalar_subclasses_and_numpy_scalars() -> None:
    """Subclasses skip the exact-type shortcut and keep the general classification."""

    class Coord(int):
        pass

    class Label(str):
        pass

    coord = Coord(7)
    label = Label("A1")
    assert as_scalar(coord) is coord
    assert as_scalar(label) is label
    assert as_scalar(XlError.NA) == XlError.NA
    assert as_scalar(np.float64(2.5)) == 2.5
    assert as_scalar(np.str_("text")) == "text"


def test_as_scalar_rejects_tuple_and_zero_row_ndarray() -> None:
    assert as_scalar((1, 2)) == XlError.VALUE
    assert as_scalar(np.empty((0, 2), dtype=object)) == XlError.VALUE


def test_to_number_rejects_range() -> None:
    def resolve(address: str) -> CellValue:
        return 1

    rng = cast(CellValue, Range("S", 1, 1, 2, 1, resolve))
    assert to_number(rng) == XlError.VALUE


def test_to_string_does_not_leak_range_repr() -> None:
    def resolve(address: str) -> CellValue:
        return 1

    text = to_string(cast(CellValue, Range("S", 1, 1, 2, 1, resolve)))
    assert text == XlError.VALUE.value
    assert "Range(" not in text


def test_get_error_walks_lazy_range_for_full_scan_precheck() -> None:
    """Generic-function precheck walks Range; lookups skip get_error instead."""
    calls: list[str] = []
    values: dict[str, CellValue] = {"S!A1": 1, "S!A2": XlError.DIV}

    def resolve(address: str) -> CellValue:
        calls.append(address)
        return values[address]

    rng = Range("S", 1, 1, 2, 1, resolve)
    assert get_error(rng) == XlError.DIV
    assert calls == ["S!A1", "S!A2"]


def test_abs_over_multi_cell_range_is_value_without_sibling_eval() -> None:
    """Non-Grid consumers reject multi-cell ranges at resolve (#VALUE!)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", "=1/0", None),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", "=ABS(S!A1:S!A3)", None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": XlError.VALUE}
        assert "S!A1" not in ev._cache
        assert "S!A2" not in ev._cache
        assert "S!A3" not in ev._cache
    assert "S!A1" not in seen
    assert "S!A2" not in seen
    assert "S!A3" not in seen


def test_text_over_multi_cell_range_is_value_not_repr() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", '=TEXT(S!A1:S!A2,"0")', None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": XlError.VALUE}


def test_index_cells_rejects_non_scalar_row_col() -> None:
    table = [[1, 2], [3, 4]]
    assert index_cells(table, [[1]], None) == XlError.VALUE
    assert index_cells(table, 1, [[2]]) == XlError.VALUE

    def resolve(address: str) -> CellValue:
        return 1

    rng = Range("S", 1, 1, 2, 1, resolve)
    assert index_cells(table, rng, None) == XlError.VALUE


def test_export_xl_index_raises_on_non_scalar_row() -> None:
    try:
        export_xl_index([[1, 2], [3, 4]], [[1]], None)
    except XlErrorException as exc:
        assert exc.code == XlError.VALUE
    else:
        raise AssertionError("expected XlErrorException")
