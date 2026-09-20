"""Workbook-backed series-graph visualization (integration).

Builds a two-series `.xlsx` with bindings, extracts a dependency graph, and
checks that `to_series_graph` plus GraphViz/Mermaid exporters collapse cell
edges to one node per series.
"""

from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import create_dependency_graph, to_graphviz, to_mermaid
from excel_grapher.exporter import to_series_graph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import bindings_document, series_entry


def _make_two_series_xlsx(path: Path) -> None:
    wb = fastpyxl.Workbook()
    inputs = wb.active
    inputs.title = "Inputs"
    inputs["A1"].value = 1
    inputs["B1"].value = 2
    inputs["A2"].value = 2
    inputs["B2"].value = 3
    calc = wb.create_sheet("Calc")
    calc["A1"].value = 1
    calc["B1"].value = 2
    calc["A2"].value = "=Inputs!A2+Inputs!B2"
    calc["B2"].value = "=A2*2"
    outputs = wb.create_sheet("Outputs")
    outputs["A1"].value = "=Calc!A2"
    wb.save(path)
    wb.close()


def _two_series_bindings() -> dict:
    return bindings_document(
        series_entry("inputs", "Inputs!A2:B2", layout="series", direction="input", header_row=1),
        series_entry("calc", "Calc!A2:B2", layout="series", direction="internal", header_row=1),
        series_entry("outputs", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_workbook_series_graph_consolidates_between_series_edges(tmp_path: Path) -> None:
    excel_path = tmp_path / "two_series.xlsx"
    _make_two_series_xlsx(excel_path)
    bindings = validate_bindings_document(_two_series_bindings())
    graph = create_dependency_graph(
        excel_path, all_series_targets(bindings, workbook=excel_path), load_values=False
    )

    series_graph = to_series_graph(graph, bindings, workbook=excel_path)

    assert [node.series_id for node in series_graph.nodes] == ["inputs", "calc", "outputs"]
    edges = {(edge.source, edge.target): edge for edge in series_graph.edges}
    assert set(edges) == {("calc", "inputs"), ("outputs", "calc")}
    assert edges[("calc", "inputs")].edge_count == 2
    assert ("calc", "calc") not in edges

    dot = to_graphviz(series_graph, rankdir="LR")
    assert '"calc" -> "inputs"' in dot
    assert "Calc!A1" not in dot

    mm = to_mermaid(series_graph)
    assert "calc" in mm
    assert "inputs" in mm
    assert "Calc!A1" not in mm
