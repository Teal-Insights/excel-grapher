"""Preserve axis intent on Node string→AST entry points (#556)."""

from __future__ import annotations

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    BinaryOpNode,
    CellRefNode,
    RelativeAxis,
    parse,
    parse_preserving_axes,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node, make_cell_node


def test_node_constructs_from_formula_ast_without_normalized_formula() -> None:
    ast = parse_preserving_axes("=A1+1", anchor="Sheet1!B1")
    node = Node(
        sheet="Sheet1",
        column="B",
        row=1,
        is_leaf=False,
        formula_ast=ast,
    )
    assert node.formula_ast is ast
    assert node.normalized_formula == "=Sheet1!A1+1"
    assert node._unparseable_formula is None


def test_node_bootstrap_preserves_relative_and_absolute_axes() -> None:
    node = Node(
        sheet="Sheet1",
        column="B",
        row=2,
        is_leaf=False,
        normalized_formula="=A1+$B$1",
    )
    expected = parse_preserving_axes("=A1+$B$1", anchor="Sheet1!B2")
    assert node.formula_ast == expected
    assert isinstance(node.formula_ast, BinaryOpNode)
    left = node.formula_ast.left
    right = node.formula_ast.right
    assert isinstance(left, CellRefNode)
    assert isinstance(right, CellRefNode)
    assert isinstance(left.ref.col, RelativeAxis)
    assert left.ref.col.offset == -1
    assert isinstance(left.ref.row, RelativeAxis)
    assert left.ref.row.offset == -1
    assert isinstance(right.ref.col, AbsoluteAxis)
    assert isinstance(right.ref.row, AbsoluteAxis)
    assert node.normalized_formula == "=Sheet1!A1+Sheet1!B1"


def test_node_bootstrap_absolute_dollar_markers() -> None:
    node = make_cell_node(
        "Sheet1",
        "C",
        3,
        is_leaf=False,
        normalized_formula="=$A$1",
    )
    assert node.formula_ast == parse_preserving_axes("=$A$1", anchor="Sheet1!C3")
    assert isinstance(node.formula_ast, CellRefNode)
    assert isinstance(node.formula_ast.ref.col, AbsoluteAxis)
    assert isinstance(node.formula_ast.ref.row, AbsoluteAxis)


def test_apply_formula_text_preserves_axes_with_host_cell() -> None:
    node = make_cell_node("Sheet1", "B", 2, is_leaf=False)
    node.apply_formula_text("=A1+$A1+A$1")
    assert node.formula_ast == parse_preserving_axes("=A1+$A1+A$1", anchor="Sheet1!B2")
    assert node.normalized_formula == "=Sheet1!A1+Sheet1!A1+Sheet1!A1"
    assert node._unparseable_formula is None


def test_apply_formula_text_does_not_absolutize_bare_a1() -> None:
    node = make_cell_node("Sheet1", "B", 1, is_leaf=False)
    node.apply_formula_text("=A1")
    assert node.formula_ast == parse_preserving_axes("=A1", anchor="Sheet1!B1")
    assert node.formula_ast != parse("=Sheet1!A1")
    assert isinstance(node.formula_ast, CellRefNode)
    assert isinstance(node.formula_ast.ref.col, RelativeAxis)
    assert node.formula_ast.ref.col.offset == -1


def test_apply_formula_text_keeps_unparseable_fallback() -> None:
    node = make_cell_node("Sheet1", "B", 1, is_leaf=False)
    text = "=SUM(IF(@A1:A3>0,1,0))"
    node.apply_formula_text(text)
    assert node.formula_ast is None
    assert node.normalized_formula == text
    assert node.has_formula


def test_identity_transit_string_fallback_preserves_mixed_axes() -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "C", 1, value=1, is_leaf=True))
    graph.add_node(make_cell_node("Sheet1", "B", 1, is_leaf=False, formula_ast=parse("=Sheet1!C1")))
    dependent = make_cell_node("Sheet1", "A", 1, is_leaf=False)
    fallback = "=Sheet1!B1+$A$1"
    dependent.formula_ast = None
    dependent._unparseable_formula = fallback
    graph.add_node(dependent)
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!C1",
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )
    start = fallback.index("Sheet1!B1")
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=DependencyCause.direct_ref,
            direct_sites_normalized=((start, start + len("Sheet1!B1")),),
        ),
    )

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.formula_ast == parse_preserving_axes("=Sheet1!C1+$A$1", anchor="Sheet1!A1")
    assert node.normalized_formula == "=Sheet1!C1+Sheet1!A1"
    assert isinstance(node.formula_ast, BinaryOpNode)
    left = node.formula_ast.left
    right = node.formula_ast.right
    assert isinstance(left, CellRefNode)
    assert isinstance(right, CellRefNode)
    assert isinstance(left.ref.col, RelativeAxis)
    assert isinstance(right.ref.col, AbsoluteAxis)
    assert isinstance(right.ref.row, AbsoluteAxis)
