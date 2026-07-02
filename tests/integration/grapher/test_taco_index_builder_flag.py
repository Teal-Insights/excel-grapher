"""Integration tests for optional TACO index on create_dependency_graph."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import TacoIndex, build_taco_index
from tests.unit.grapher.range_compression.parity_helpers import assert_taco_parity


def test_create_dependency_graph_default_has_no_taco_index(tmp_path: Path) -> None:
    path = tmp_path / "chain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1.0)
    ws.write_formula(0, 1, "=A1")
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!B1"], load_values=False)
    assert graph.taco_index is None


def test_create_dependency_graph_taco_index_flag_attaches_index(tmp_path: Path) -> None:
    path = tmp_path / "rr.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    for row in range(3, 8):
        ws.write_number(row - 1, 1, float(row))
        ws.write_formula(row - 1, 3, f"=B{row}")
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Data!D3:D7"],
        load_values=False,
        taco_index=True,
    )
    assert graph.taco_index is not None
    assert isinstance(graph.taco_index, TacoIndex)
    assert len(graph.taco_index.compressed_edges) >= 1
    assert_taco_parity(graph, graph.taco_index)


def test_taco_index_flag_matches_explicit_build(tmp_path: Path) -> None:
    path = tmp_path / "rr.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    for row in range(3, 8):
        ws.write_number(row - 1, 1, float(row))
        ws.write_formula(row - 1, 3, f"=B{row}")
    wb.close()

    attached = create_dependency_graph(path, ["Data!D3:D7"], load_values=False, taco_index=True)
    manual = create_dependency_graph(path, ["Data!D3:D7"], load_values=False)
    manual_index = build_taco_index(manual)

    assert attached.taco_index is not None
    assert len(attached.taco_index.compressed_edges) == len(manual_index.compressed_edges)
    assert len(attached.taco_index.single_edges) == len(manual_index.single_edges)
