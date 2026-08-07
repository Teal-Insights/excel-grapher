"""Graphviz export reflects workbook-backed dependency graphs (integration).

Builds small `.xlsx` files, runs `create_dependency_graph` and `to_graphviz`,
and asserts DOT output includes expected nodes, edges, and shape hints.
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import create_dependency_graph, to_graphviz
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node


def _make_chain_xlsx(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 2
    ws["A2"].value = 3
    ws["A3"].value = "=A1+A2"
    ws["A4"].value = "=A3*2"
    wb.save(path)
    wb.close()


def test_to_graphviz_contains_nodes_edges_and_shapes(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _make_chain_xlsx(excel_path)

    graph = create_dependency_graph(
        excel_path, ["Sheet1!A4"], load_values=False, store_raw_formula=True
    )
    dot = to_graphviz(graph, rankdir="LR")

    assert "digraph dependencies" in dot
    assert "rankdir=LR" in dot

    # Nodes exist
    assert '"Sheet1!A1"' in dot
    assert '"Sheet1!A4"' in dot

    # Edges exist (A4 depends on A3; A3 depends on A1 and A2)
    assert '"Sheet1!A4" -> "Sheet1!A3"' in dot
    assert '"Sheet1!A3" -> "Sheet1!A1"' in dot
    assert '"Sheet1!A3" -> "Sheet1!A2"' in dot

    # Leaf nodes are boxes; formula nodes are ellipses (labels include formula by default)
    assert '"Sheet1!A1" [label="Sheet1!A1" shape=box' in dot
    assert '"Sheet1!A4" [label="Sheet1!A4\\n=A3*2" shape=ellipse' in dot


def test_to_graphviz_can_omit_formula_labels(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _make_chain_xlsx(excel_path)
    graph = create_dependency_graph(
        excel_path, ["Sheet1!A4"], load_values=False, store_raw_formula=True
    )
    dot = to_graphviz(graph, rankdir="LR", include_formula_on_nodes=False)
    assert '"Sheet1!A4" [label="Sheet1!A4" shape=ellipse' in dot


def test_to_graphviz_truncates_formula(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _make_chain_xlsx(excel_path)
    graph = create_dependency_graph(
        excel_path, ["Sheet1!A4"], load_values=False, store_raw_formula=True
    )
    dot = to_graphviz(graph, rankdir="LR", max_formula_length=4)
    assert '"Sheet1!A4" [label="Sheet1!A4\\n=A3*..." shape=ellipse' in dot


def test_to_graphviz_invalid_max_formula_length() -> None:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.node import Node

    g = DependencyGraph()
    g.add_node(
        Node(
            sheet="S",
            column="A",
            row=1,
            formula=None,
            normalized_formula=None,
            value=1,
            is_leaf=True,
        )
    )
    with pytest.raises(ValueError, match="max_formula_length"):
        to_graphviz(g, max_formula_length=0)


def test_to_graphviz_uses_graph_sheet_order_for_node_listing() -> None:
    g = DependencyGraph(sheet_order=["Later", "Earlier"])
    g.add_node(Node("Earlier", "A", 1, None, None, 1, True))
    g.add_node(Node("Later", "A", 1, None, None, 1, True))

    dot = to_graphviz(g)
    later_idx = dot.index('"Later!A1" [')
    earlier_idx = dot.index('"Earlier!A1" [')
    assert later_idx < earlier_idx
