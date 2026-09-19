"""Workbook-backed sheet-graph visualization (integration).

Builds a two-sheet `.xlsx`, extracts a dependency graph, and checks that
`to_sheet_graph` plus GraphViz/Mermaid exporters collapse cell edges to one
node per sheet.
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import create_dependency_graph, to_graphviz, to_mermaid, to_sheet_graph


def _make_two_sheet_xlsx(path: Path) -> None:
    wb = fastpyxl.Workbook()
    inputs = wb.active
    inputs.title = "Inputs"
    inputs["A1"].value = 2
    inputs["B1"].value = 3
    calc = wb.create_sheet("Calc")
    calc["A1"].value = "=Inputs!A1+Inputs!B1"
    calc["B1"].value = "=A1*2"
    outputs = wb.create_sheet("Outputs")
    outputs["A1"].value = "=Calc!A1"
    wb.save(path)
    wb.close()


def test_workbook_sheet_graph_consolidates_between_sheet_edges(tmp_path: Path) -> None:
    excel_path = tmp_path / "two_sheet.xlsx"
    _make_two_sheet_xlsx(excel_path)

    graph = create_dependency_graph(excel_path, ["Outputs!A1"], load_values=False)
    sheet_graph = to_sheet_graph(graph)

    assert [node.name for node in sheet_graph.nodes] == ["Inputs", "Calc", "Outputs"]
    edges = {(edge.source, edge.target): edge for edge in sheet_graph.edges}
    assert set(edges) == {("Calc", "Inputs"), ("Outputs", "Calc")}
    assert edges[("Calc", "Inputs")].edge_count == 2
    assert ("Calc", "Calc") not in edges

    dot = to_graphviz(sheet_graph, rankdir="LR")
    assert '"Calc" -> "Inputs"' in dot
    assert "Calc!A1" not in dot

    mm = to_mermaid(sheet_graph)
    assert "Calc" in mm
    assert "Inputs" in mm
    assert "Calc!A1" not in mm
