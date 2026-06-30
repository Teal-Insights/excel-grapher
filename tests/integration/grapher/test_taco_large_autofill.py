"""Large synthetic autofill workbook TACO parity and compression ratio."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import PatternKind, build_taco_index
from tests.unit.grapher.range_compression.parity_helpers import assert_taco_parity


def test_large_rr_column_parity_and_compression_ratio(tmp_path: Path) -> None:
    path = tmp_path / "large_rr.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    first, last = 3, 102
    for row in range(first, last + 1):
        ws.write_number(row - 1, 1, float(row))
        ws.write_number(row - 1, 2, float(row * 2))
        ws.write_formula(row - 1, 3, f"=B{row}*C{row}")
    wb.close()

    graph = create_dependency_graph(path, [f"Data!D{first}:D{last}"], load_values=False)
    index = build_taco_index(graph)
    assert_taco_parity(graph, index)

    rr_edges = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr_edges) == 2
    cell_edges = sum(
        len(graph.get_dependencies(k))
        for k in graph
        if (node := graph.get_node(k)) is not None and not node.is_leaf
    )
    assert len(index.compressed_edges) < cell_edges // 10
