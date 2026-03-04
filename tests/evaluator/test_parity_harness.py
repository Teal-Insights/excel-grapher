from __future__ import annotations

import pytest

import tests.evaluator.parity_harness as parity_harness
from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.name_utils import parse_address
from tests.evaluator.parity_harness import (
    assert_code_does_not_embed_symbols,
    assert_codegen_matches_evaluator,
)


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


def test_fail_fast_reports_first_dependency_mismatch(monkeypatch: pytest.MonkeyPatch) -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 1))
    graph.add_node(_make_node("S!B1", "=S!A1+1", None))
    graph.add_node(_make_node("S!C1", "=S!B1+1", None))
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!C1", "S!B1")

    def _always_false(_a: object, _b: object, *, rtol: float, atol: float) -> bool:
        return False

    monkeypatch.setattr(parity_harness, "_values_equal", _always_false)

    with pytest.raises(AssertionError) as exc:
        assert_codegen_matches_evaluator(
            graph,
            ["S!C1"],
            dependency_order=True,
            fail_fast=True,
        )

    message = str(exc.value)
    assert "First parity mismatch" in message
    assert "S!A1" in message


def test_parity_harness_simple_arithmetic_and_pruning() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1+S!A2", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == 15.0

    # This graph does not use SUM, so xl_sum should not be embedded.
    assert_code_does_not_embed_symbols(result.generated_code, absent={"xl_sum"})


def test_parity_row_with_reference_and_offset() -> None:
    graph = _make_graph(
        _make_node("S!B3", None, 10),
        _make_node("S!C1", None, 2),
        _make_node("S!C2", None, 0),
        _make_node("S!A1", "=ROW(S!B3)", None),
        _make_node("S!A2", "=ROW(OFFSET(S!B3, S!C1, S!C2))", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2"])
    assert result.generated_results["S!A1"] == 3
    assert result.generated_results["S!A2"] == 5


def test_parity_column_and_columns_with_offset() -> None:
    graph = _make_graph(
        _make_node("S!D4", None, 10),
        _make_node("S!C1", None, 0),
        _make_node("S!C2", None, 2),
        _make_node("S!A1", "=COLUMN(S!D4)", None),
        _make_node("S!A2", "=COLUMNS(S!D4)", None),
        _make_node("S!A3", "=COLUMN(OFFSET(S!D4, S!C1, S!C2))", None),
        _make_node("S!A4", "=COLUMNS(OFFSET(S!D4, S!C1, S!C2, 1, 3))", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2", "S!A3", "S!A4"])
    assert result.generated_results["S!A1"] == 4
    assert result.generated_results["S!A2"] == 1
    assert result.generated_results["S!A3"] == 6
    assert result.generated_results["S!A4"] == 3

