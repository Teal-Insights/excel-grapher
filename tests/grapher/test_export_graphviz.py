from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import create_dependency_graph, to_graphviz


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

    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=False)
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

    # Leaf nodes are boxes; formula nodes are ellipses
    assert '"Sheet1!A1" [label="Sheet1!A1" shape=box' in dot
    assert '"Sheet1!A4" [label="Sheet1!A4" shape=ellipse' in dot
