"""Consumer migration onto `formula_ast` (#544)."""

from __future__ import annotations

from unittest.mock import patch

import excel_grapher.evaluator.evaluator as evaluator_module
from excel_grapher import FormulaEvaluator
from excel_grapher.core.formula_ast import (
    RelativeAxis,
    parse,
    parse_preserving_axes,
    unparse_normalized_formula,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.formula_label import display_formula
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node


def test_has_formula_true_for_ast_or_normalized_text() -> None:
    ast_only = make_cell_node(
        "Sheet1",
        "B",
        1,
        is_leaf=False,
        formula_ast=parse("=Sheet1!A1+1"),
    )
    assert ast_only.has_formula
    text_only = make_cell_node(
        "Sheet1",
        "C",
        1,
        is_leaf=False,
        normalized_formula="=SUM(IF(@Sheet1!A1:A3>0,1,0))",
    )
    assert text_only.has_formula
    leaf = make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True)
    assert not leaf.has_formula


def test_set_node_ast_derives_normalized_formula() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, is_leaf=True, value=1))
    graph.add_node(make_cell_node("Sheet1", "B", 1, is_leaf=False))
    ast = parse("=Sheet1!A1+2")
    graph.set_node_ast("Sheet1!B1", ast)
    view = graph.get_node("Sheet1!B1")
    assert view is not None
    assert view.formula_ast == ast
    assert view.normalized_formula == "=Sheet1!A1+2"
    assert view.formula is None

    graph.set_node_ast("Sheet1!B1", None)
    cleared = graph.get_node("Sheet1!B1")
    assert cleared is not None
    assert cleared.formula_ast is None
    assert cleared.normalized_formula is None
    assert not cleared.has_formula


def test_evaluator_uses_formula_ast_even_when_normalized_text_is_stale() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("S", "A", 1, value=10, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "S",
            "B",
            1,
            is_leaf=False,
            normalized_formula="=S!A1+999",
            formula_ast=parse("=S!A1+1"),
        )
    )
    graph.add_edge("S!B1", "S!A1")

    parse_calls = 0
    original_parse = evaluator_module.parse

    def counting_parse(formula: str):
        nonlocal parse_calls
        parse_calls += 1
        return original_parse(formula)

    with FormulaEvaluator(graph) as ev, patch.object(evaluator_module, "parse", counting_parse):
        assert ev.evaluate("S!B1") == 11.0
        assert parse_calls == 0


def test_display_formula_prefers_raw_then_unparsed_ast() -> None:
    raw = make_cell_node(
        "Sheet1",
        "B",
        1,
        formula="=A1",
        normalized_formula="=Sheet1!A1",
        formula_ast=parse("=Sheet1!A1"),
        is_leaf=False,
    )
    assert display_formula(raw) == "=A1"

    ast_only = make_cell_node(
        "Sheet1",
        "C",
        1,
        is_leaf=False,
        formula_ast=parse_preserving_axes("=A1*2", anchor="Sheet1!C1"),
    )
    assert display_formula(ast_only) == "=Sheet1!A1*2"


def _direct_edge(graph: DependencyGraph, dependent: str, precedent: str) -> None:
    dep = graph.get_node(dependent)
    assert dep is not None
    normalized = dep.normalized_formula
    assert normalized is not None
    start = normalized.index(precedent)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=DependencyCause.direct_ref,
            direct_sites_normalized=((start, start + len(precedent)),),
        ),
    )


def test_identity_transit_rewrites_ast_and_keeps_unrelated_relative_axes() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "C", 1, value=1, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=C1", anchor="Sheet1!B1"),
        )
    )
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "A",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=B1+C2", anchor="Sheet1!A1"),
        )
    )
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.formula_ast is not None
    # Unrelated C2 relative offset (col+2, row+1 from A1) is preserved.
    from excel_grapher.core.formula_ast import BinaryOpNode, CellRefNode

    assert isinstance(node.formula_ast, BinaryOpNode)
    right = node.formula_ast.right
    assert isinstance(right, CellRefNode)
    assert isinstance(right.ref.col, RelativeAxis)
    assert right.ref.col.offset == 2
    assert isinstance(right.ref.row, RelativeAxis)
    assert right.ref.row.offset == 1
    assert node.normalized_formula == unparse_normalized_formula(
        node.formula_ast, anchor="Sheet1!A1"
    )
    assert node.normalized_formula == "=Sheet1!C1+Sheet1!C2"


def test_identity_transit_derives_a1_from_ast_not_string_replace() -> None:
    """Range endpoints are not identity sites; derived A1 must follow the AST."""
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "C", 1, value=1, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            is_leaf=False,
            formula_ast=parse("=Sheet1!C1"),
        )
    )
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "A",
            1,
            is_leaf=False,
            formula_ast=parse("=SUM(Sheet1!B1:B3)+Sheet1!B1"),
        )
    )
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.formula_ast is not None
    assert node.formula_ast == parse("=SUM(Sheet1!B1:B3)+Sheet1!C1")
    assert node.normalized_formula == unparse_normalized_formula(
        node.formula_ast, anchor="Sheet1!A1"
    )
    assert node.normalized_formula == "=SUM(Sheet1!B1:B3)+Sheet1!C1"


def test_structural_inline_derives_a1_from_spliced_ast() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "D", 1, value=5, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            is_leaf=False,
            formula_ast=parse("=Sheet1!D1*2"),
        )
    )
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "A",
            1,
            is_leaf=False,
            formula_ast=parse("=Sheet1!B1+1"),
        )
    )
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.formula_ast is not None
    assert node.formula_ast == parse("=Sheet1!D1*2+1")
    assert node.normalized_formula == unparse_normalized_formula(
        node.formula_ast, anchor="Sheet1!A1"
    )
    assert node.normalized_formula == "=Sheet1!D1*2+1"
