"""Integration tests for building a TACO index from a dependency graph."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import TacoIndex, build_taco_index
from tests.unit.grapher.range_compression.parity_helpers import assert_taco_parity


def test_build_taco_index_returns_index(tmp_path: Path) -> None:
    path = tmp_path / "rr.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    for row in range(3, 8):
        ws.write_number(row - 1, 1, float(row))
        ws.write_formula(row - 1, 3, f"=B{row}")
    wb.close()

    graph = create_dependency_graph(path, ["Data!D3:D7"], load_values=False, store_raw_formula=True)
    index = build_taco_index(graph)
    assert isinstance(index, TacoIndex)
    assert len(index.compressed_edges) >= 1
    assert_taco_parity(graph, index)


def test_build_taco_index_is_deterministic(tmp_path: Path) -> None:
    path = tmp_path / "rr.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    for row in range(3, 8):
        ws.write_number(row - 1, 1, float(row))
        ws.write_formula(row - 1, 3, f"=B{row}")
    wb.close()

    graph = create_dependency_graph(path, ["Data!D3:D7"], load_values=False, store_raw_formula=True)
    first = build_taco_index(graph)
    second = build_taco_index(graph)

    assert len(first.compressed_edges) == len(second.compressed_edges)
    assert len(first.single_edges) == len(second.single_edges)
