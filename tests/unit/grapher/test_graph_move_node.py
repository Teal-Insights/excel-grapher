"""`DependencyGraph.move_node` preserves resolved formula targets (#549)."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import CellKey
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    BinaryOpNode,
    CellRefNode,
    RelativeAxis,
    parse_preserving_axes,
    resolve_cell_ref,
)
from excel_grapher.grapher.formula_shapes import warm_formula_shapes
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node


def _add_leaf(graph: DependencyGraph, sheet: str, column: str, row: int, value: object = 1) -> None:
    graph.add_node(make_cell_node(sheet, column, row, value=value, is_leaf=True))


def _add_formula(
    graph: DependencyGraph,
    sheet: str,
    column: str,
    row: int,
    formula: str,
    *,
    deps: tuple[str, ...] = (),
) -> None:
    anchor = f"{sheet}!{column}{row}"
    ast = parse_preserving_axes(formula, anchor=anchor)
    graph.add_node(
        make_cell_node(
            sheet,
            column,
            row,
            formula=formula,
            formula_ast=ast,
            is_leaf=not deps,
        )
    )
    for dep in deps:
        graph.add_edge(anchor, dep)


def test_move_node_host_keeps_resolved_targets() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 1)
    _add_formula(graph, "Sheet1", "B", 2, "=A1", deps=("Sheet1!A1",))

    graph.move_node("Sheet1!B2", "Sheet1!C3")

    assert "Sheet1!B2" not in graph
    view = graph.get_node("Sheet1!C3")
    assert view is not None
    assert view.key == "Sheet1!C3"
    assert view.normalized_formula == "=Sheet1!A1"
    assert isinstance(view.formula_ast, CellRefNode)
    assert view.formula_ast.ref.col == RelativeAxis(-2)
    assert view.formula_ast.ref.row == RelativeAxis(-2)
    assert resolve_cell_ref(view.formula_ast.ref, view.address) == "Sheet1!A1"
    assert graph.get_dependencies("Sheet1!C3") == frozenset({"Sheet1!A1"})
    assert graph.get_dependents("Sheet1!A1") == frozenset({"Sheet1!C3"})


def test_move_node_rewrites_dependent_refs() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 1)
    _add_formula(graph, "Sheet1", "B", 2, "=A1", deps=("Sheet1!A1",))

    graph.move_node("Sheet1!A1", "Sheet1!C3")

    host = graph.get_node("Sheet1!B2")
    assert host is not None
    assert host.normalized_formula == "=Sheet1!C3"
    assert isinstance(host.formula_ast, CellRefNode)
    assert host.formula_ast.ref.col == RelativeAxis(1)
    assert host.formula_ast.ref.row == RelativeAxis(1)
    assert resolve_cell_ref(host.formula_ast.ref, host.address) == "Sheet1!C3"
    assert graph.get_dependencies("Sheet1!B2") == frozenset({"Sheet1!C3"})
    assert graph.get_dependents("Sheet1!C3") == frozenset({"Sheet1!B2"})
    assert "Sheet1!A1" not in graph


def test_move_node_mixed_axes_range_and_whole_leaves() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 1)
    _add_leaf(graph, "Sheet1", "A", 3)
    _add_formula(graph, "Sheet1", "B", 2, "=$A1+SUM(A1:A3)", deps=("Sheet1!A1", "Sheet1!A3"))
    _add_formula(graph, "Sheet1", "D", 1, "=A:A", deps=("Sheet1!A1", "Sheet1!A3"))

    graph.move_node("Sheet1!B2", "Sheet1!C4")

    moved = graph.get_node("Sheet1!C4")
    assert moved is not None
    assert moved.normalized_formula == "=Sheet1!A1+SUM(Sheet1!A1:A3)"
    assert isinstance(moved.formula_ast, BinaryOpNode)
    dollar_a = moved.formula_ast.left
    assert isinstance(dollar_a, CellRefNode)
    assert dollar_a.ref.col == AbsoluteAxis(1)
    assert dollar_a.ref.row == RelativeAxis(-3)

    graph.move_node("Sheet1!A1", "Sheet1!B5")
    dependent = graph.get_node("Sheet1!C4")
    assert dependent is not None
    assert dependent.normalized_formula == "=Sheet1!B5+SUM(Sheet1!B5:A3)"
    whole = graph.get_node("Sheet1!D1")
    assert whole is not None
    assert whole.normalized_formula == "=Sheet1!A:A"


def test_move_node_reinterns_formula_shapes() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 2)
    _add_leaf(graph, "Sheet1", "A", 3)
    _add_formula(graph, "Sheet1", "B", 2, "=A2", deps=("Sheet1!A2",))
    _add_formula(graph, "Sheet1", "B", 3, "=A3", deps=("Sheet1!A3",))
    graph.formula_shapes = warm_formula_shapes(graph)

    before_b2 = graph.formula_shapes.lookup("Sheet1!B2")
    before_b3 = graph.formula_shapes.lookup("Sheet1!B3")
    assert before_b2 is not None and before_b3 is not None
    assert before_b2[0] == before_b3[0]
    assert before_b2[2] == before_b3[2]

    graph.move_node("Sheet1!B2", "Sheet1!C4")

    table = graph.formula_shapes
    assert table is not None
    assert table.lookup("Sheet1!B2") is None
    after_moved = table.lookup("Sheet1!C4")
    after_b3 = table.lookup("Sheet1!B3")
    assert after_moved is not None and after_b3 is not None
    assert after_moved[0] == after_b3[0]
    assert after_moved[2] != after_b3[2]
    assert after_b3[2] == before_b3[2]


def test_move_node_fail_closed_on_direct_address_assignment() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 1)
    node = graph._get_internal_node("Sheet1!A1")
    assert node is not None
    with pytest.raises(ValueError, match="move_node"):
        node.address = CellKey("Sheet1!Z9")
    assert node.key == "Sheet1!A1"
    assert graph.get_node("Sheet1!A1") is not None


def test_move_node_rejects_missing_and_occupied_and_non_cell() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 1)
    _add_leaf(graph, "Sheet1", "B", 1)
    with pytest.raises(KeyError, match="Sheet1!Z9"):
        graph.move_node("Sheet1!Z9", "Sheet1!C1")
    with pytest.raises(ValueError, match="already exists"):
        graph.move_node("Sheet1!A1", "Sheet1!B1")
    with pytest.raises(ValueError, match="single cell"):
        graph.move_node("Sheet1!A1", "Sheet1!C1:D2")


def test_move_node_same_key_is_noop() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 1)
    _add_formula(graph, "Sheet1", "B", 1, "=A1", deps=("Sheet1!A1",))
    ast_before = graph.get_node("Sheet1!B1")
    assert ast_before is not None
    formula_ast = ast_before.formula_ast
    graph.move_node("Sheet1!B1", "Sheet1!B1")
    after = graph.get_node("Sheet1!B1")
    assert after is not None
    assert after.formula_ast is formula_ast


def test_move_node_raises_on_unparseable_dependent() -> None:
    graph = DependencyGraph()
    _add_leaf(graph, "Sheet1", "A", 1)
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            normalized_formula="=SUM(IF(@Sheet1!A1>0,1,0))",
            is_leaf=False,
        )
    )
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    with pytest.raises(ValueError, match="unparseable"):
        graph.move_node("Sheet1!A1", "Sheet1!C3")
    assert "Sheet1!A1" in graph
