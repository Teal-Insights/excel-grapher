"""Integration tests for projected visualization export."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher.exporter import IdentityTransitCompression, to_web_viz_payload
from excel_grapher.grapher import create_dependency_graph, to_networkx
from excel_grapher.grapher.lightweight_viz import lightweight_viz_flat

pytest.importorskip("networkx")


def test_projected_networkx_omits_transit_nodes_without_mutating_graph(tmp_path: Path) -> None:
    workbook_path = tmp_path / "identity_target.xlsx"
    wb = xlsxwriter.Workbook(workbook_path)
    ws = wb.add_worksheet("Engine")
    ws.write_number("C6", 10)
    out = wb.add_worksheet("Outputs")
    out.write_formula("B12", "=Engine!C6")
    out.write_formula("B14", "=Outputs!B12+1")
    wb.close()

    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B14"],
        capture_dependency_provenance=True,
    )
    original_node_count = len(graph)

    projection = IdentityTransitCompression().project(graph)
    nx_graph = to_networkx(projection)
    payload = to_web_viz_payload(nx_graph)
    flat = lightweight_viz_flat(payload)

    assert len(graph) == original_node_count
    assert "Outputs!B12" in graph
    assert "Outputs!B12" not in projection
    assert "Outputs!B12" not in nx_graph
    assert flat.stats.node_count < original_node_count
