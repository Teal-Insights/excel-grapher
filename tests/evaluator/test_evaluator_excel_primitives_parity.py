"""
Focused regression tests for Excel primitives that often diverge in the evaluator.

Pairs with small synthetic graphs; use alongside workbook triage tests.
"""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.name_utils import parse_address
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


def test_evaluator_row_without_reference_uses_calling_cell_row() -> None:
    graph = _make_graph(_make_node("S!B5", "=ROW()", None))
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B5"]) == {"S!B5": 5}


def test_evaluator_column_without_reference_uses_calling_cell_column() -> None:
    graph = _make_graph(_make_node("S!C4", "=COLUMN()", None))
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!C4"]) == {"S!C4": 3}


def test_codegen_matches_evaluator_row_without_reference() -> None:
    graph = _make_graph(_make_node("S!D9", "=ROW()", None))
    assert_codegen_matches_evaluator(graph, ["S!D9"])


def test_codegen_matches_evaluator_column_without_reference() -> None:
    graph = _make_graph(_make_node("S!F2", "=COLUMN()", None))
    assert_codegen_matches_evaluator(graph, ["S!F2"])
