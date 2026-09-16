from __future__ import annotations

import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.types import XlError
from tests.integration.utils import parity_harness as harness
from tests.integration.utils.parity_harness import evaluate_targets


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


def test_evaluator_simple_arithmetic() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1+S!A2", None),
    )
    results = evaluate_targets(graph, ["S!B1"])
    assert results["S!B1"] == 15.0


def test_evaluator_row_with_reference_and_offset() -> None:
    graph = _make_graph(
        _make_node("S!B3", None, 10),
        _make_node("S!C1", None, 2),
        _make_node("S!C2", None, 0),
        _make_node("S!A1", "=ROW(S!B3)", None),
        _make_node("S!A2", "=ROW(OFFSET(S!B3, S!C1, S!C2))", None),
    )
    with FormulaEvaluator(graph) as ev:
        results = ev.evaluate(["S!A1", "S!A2"])
    assert results["S!A1"] == 3
    assert results["S!A2"] == 5


def test_evaluator_row_without_reference() -> None:
    graph = _make_graph(_make_node("S!E12", "=ROW()", None))
    results = evaluate_targets(graph, ["S!E12"])
    assert results["S!E12"] == 12


def test_evaluator_column_and_columns_with_offset() -> None:
    graph = _make_graph(
        _make_node("S!D4", None, 10),
        _make_node("S!C1", None, 0),
        _make_node("S!C2", None, 2),
        _make_node("S!A1", "=COLUMN(S!D4)", None),
        _make_node("S!A2", "=COLUMNS(S!D4)", None),
        _make_node("S!A3", "=COLUMN(OFFSET(S!D4, S!C1, S!C2))", None),
        _make_node("S!A4", "=COLUMNS(OFFSET(S!D4, S!C1, S!C2, 1, 3))", None),
    )
    with FormulaEvaluator(graph) as ev:
        results = ev.evaluate(["S!A1", "S!A2", "S!A3", "S!A4"])
    assert results["S!A1"] == 4
    assert results["S!A2"] == 1
    assert results["S!A3"] == 6
    assert results["S!A4"] == 3


def test_evaluate_targets_routes_through_inverted_tree_export(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Parity harness must emit a bindings-backed inverted-tree package (#863 A1)."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1+S!A2", None),
    )
    original = harness.generate_inverted_tree_modules
    calls: list[object] = []

    def wrapped(*args: object, **kwargs: object) -> dict[str, str]:
        calls.append(kwargs or args)
        return original(*args, **kwargs)

    monkeypatch.setattr(harness, "generate_inverted_tree_modules", wrapped)
    results = evaluate_targets(graph, ["S!B1"])
    assert calls, "evaluate_targets must call generate_inverted_tree_modules"
    assert results["S!B1"] == 15.0


def test_evaluate_targets_quoted_sheet_names() -> None:
    graph = _make_graph(
        _make_node("Product Lookup!A1", None, 10),
        _make_node("Product Lookup!B1", "='Product Lookup'!A1+5", None),
    )
    results = evaluate_targets(graph, ["Product Lookup!B1"])
    assert results["Product Lookup!B1"] == 15.0


def test_evaluate_targets_maps_export_errors_to_evaluator_sentinels() -> None:
    graph = _make_graph(_make_node("S!A1", "=1/0", None))
    results = evaluate_targets(graph, ["S!A1"])
    assert results["S!A1"] == XlError.DIV
