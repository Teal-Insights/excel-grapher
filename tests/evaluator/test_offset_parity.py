"""Parity tests for OFFSET (evaluator vs generated runtime)."""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.name_utils import parse_address
from excel_grapher.evaluator.types import XlError
from tests.evaluator.parity_harness import assert_codegen_matches_evaluator


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


def test_offset_parity_static_single_cell() -> None:
    """OFFSET with constant offsets resolves to single cell; evaluator and generated code agree."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!B1", "=OFFSET(S!A1, 1, 0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == 20


def test_offset_parity_static_range_sum() -> None:
    """OFFSET with constant size consumed by SUM; both runtimes produce same total."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!C1", "=SUM(OFFSET(S!A1, 0, 0, 2, 2))", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.generated_results["S!C1"] == 33.0


def test_offset_parity_dynamic_row_offset() -> None:
    """OFFSET with row offset from cell (dynamic); evaluator and generated code agree."""
    graph = _make_graph(
        _make_node("S!A1", None, 5),
        _make_node("S!A2", None, 10),
        _make_node("S!A3", None, 15),
        _make_node("S!B1", None, 1),
        _make_node("S!C1", "=OFFSET(S!A1, S!B1, 0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.generated_results["S!C1"] == 10


def test_offset_parity_negative_offset() -> None:
    """OFFSET with negative row offset; both runtimes return same value."""
    graph = _make_graph(
        _make_node("S!A1", None, 100),
        _make_node("S!A2", None, 200),
        _make_node("S!A3", None, 300),
        _make_node("S!B1", "=OFFSET(S!A3, -2, 0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == 100


def test_offset_parity_invalid_returns_ref_error() -> None:
    """OFFSET that resolves to invalid reference (e.g. row 0) returns #REF! in both runtimes."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", "=OFFSET(S!A1, -1, 0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == XlError.REF
