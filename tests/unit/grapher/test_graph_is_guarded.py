"""`DependencyGraph.is_guarded` and interned edge storage (#491)."""

from __future__ import annotations

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef, Compare, Literal
from excel_grapher.grapher.node import make_cell_node


def _graph_with_shared_condition() -> DependencyGraph:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True))
    graph.add_node(make_cell_node("Sheet1", "B", 1, formula="=IF(A1=1,1,0)", is_leaf=False))
    graph.add_node(make_cell_node("Sheet1", "C", 1, formula="=IF(A1=1,2,0)", is_leaf=False))
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!A1",
        guard=Compare(CellRef("Sheet1!A1"), "=", Literal(1)),
    )
    graph.add_edge(
        "Sheet1!C1",
        "Sheet1!A1",
        guard=Compare(CellRef("Sheet1!A1"), "=", Literal(1)),
    )
    return graph


def test_is_guarded_is_true_only_for_guarded_edges() -> None:
    graph = _graph_with_shared_condition()
    graph.add_edge("Sheet1!B1", "Sheet1!C1")  # unconditional
    assert graph.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert graph.is_guarded("Sheet1!C1", "Sheet1!A1")
    assert not graph.is_guarded("Sheet1!B1", "Sheet1!C1")
    assert not graph.is_guarded("Sheet1!A1", "Sheet1!B1")


def test_is_guarded_normalizes_keys() -> None:
    graph = _graph_with_shared_condition()
    assert graph.is_guarded("'Sheet1'!B1", "'Sheet1'!A1")


def test_add_edge_interns_identical_guards_across_edges() -> None:
    graph = _graph_with_shared_condition()
    g_b = graph.get_edge_guard("Sheet1!B1", "Sheet1!A1")
    g_c = graph.get_edge_guard("Sheet1!C1", "Sheet1!A1")
    assert g_b is not None and g_c is not None
    assert g_b is g_c
