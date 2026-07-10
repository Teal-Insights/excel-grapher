"""Smoke tests for hand-built mixed cell/row graphs (issue #374 sprint 4)."""

from __future__ import annotations

from excel_grapher.grapher.export import to_mermaid
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node, NodeKind, make_row_node
from excel_grapher.grapher.range_compression.grouping import column_adjacent_groups
from excel_grapher.grapher.subgraph import select_path_induced_subgraph


def _formula_cell(sheet: str, col: str, row: int, formula: str) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=None,
        is_leaf=False,
        is_target=True,
    )


def _leaf_cell(sheet: str, col: str, row: int, value: object = 0) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=None,
        normalized_formula=None,
        value=value,
        is_leaf=True,
    )


def test_smoke_cell_formula_depends_on_row_precedent() -> None:
    g = DependencyGraph()
    row = make_row_node("Sheet1", 63, "D", "Y")
    cell = _formula_cell("Sheet1", "A", 63, "=SUM(D63:Y63)")
    g.add_node(row)
    g.add_node(cell)
    g.add_edge(cell.key, row.key)

    assert g.get_dependencies(cell.key) == frozenset({row.key})
    assert g.get_dependents(row.key) == frozenset({cell.key})
    assert g.evaluation_order() == [row.key, cell.key]
    assert g.leaf_keys() == [row.key]
    assert g.formula_keys() == [cell.key]
    assert g.target_keys() == [cell.key]

    view = g.get_node(row.key)
    assert view is not None
    assert view.kind is NodeKind.row


def test_smoke_path_induced_subgraph_preserves_row_nodes() -> None:
    g = DependencyGraph()
    row = make_row_node("Sheet1", 63, "D", "Y")
    cell = _formula_cell("Sheet1", "A", 63, "=SUM(D63:Y63)")
    unrelated = _leaf_cell("Sheet1", "Z", 99)
    g.add_node(row)
    g.add_node(cell)
    g.add_node(unrelated)
    g.add_edge(cell.key, row.key)

    # Edge direction: dependent -> precedent, so source=cell, target=row.
    sub = select_path_induced_subgraph(
        g,
        source_keys=[cell.key],
        target_keys=[row.key],
    )
    assert row.key in sub
    assert cell.key in sub
    assert unrelated.key not in sub
    view = sub.get_node(row.key)
    assert view is not None
    assert view.kind is NodeKind.row
    assert view.min_col == "D"
    assert view.max_col == "Y"
    assert sub.get_dependencies(cell.key) == frozenset({row.key})


def test_smoke_taco_grouping_skips_row_nodes() -> None:
    g = DependencyGraph()
    row = make_row_node("Sheet1", 63, "D", "Y", formula="=1", is_leaf=False)
    a1 = _formula_cell("Sheet1", "A", 1, "=1")
    a2 = _formula_cell("Sheet1", "A", 2, "=1")
    g.add_node(row)
    g.add_node(a1)
    g.add_node(a2)

    groups = column_adjacent_groups(g, min_len=2)
    flat = {key for group in groups for key in group}
    assert row.key not in flat
    assert flat == {a1.key, a2.key}


def test_smoke_mermaid_export_with_row_node() -> None:
    g = DependencyGraph()
    row = make_row_node("Sheet1", 63, "D", "Y")
    cell = _formula_cell("Sheet1", "A", 63, "=SUM(D63:Y63)")
    g.add_node(row)
    g.add_node(cell)
    g.add_edge(cell.key, row.key)

    text = to_mermaid(g)
    assert "Sheet1!A63" in text
    assert "Sheet1!D63:Y63" in text
