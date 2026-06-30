"""Tests for TACO RR-Chain pattern compression."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import PatternKind, build_taco_index

from .parity_helpers import assert_taco_parity


def _make_node(key: str, formula: str | None, *, is_leaf: bool = False) -> Node:
    sheet, rest = key.split("!", 1)
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=1 if is_leaf else None,
        is_leaf=is_leaf,
    )


def test_rr_chain_manual_graph() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("Sheet1!P3", formula=None, is_leaf=True))
    for row in range(4, 8):
        graph.add_node(_make_node(f"Sheet1!P{row}", formula=f"=P{row - 1}+1"))
        graph.add_edge(
            f"Sheet1!P{row}",
            f"Sheet1!P{row - 1}",
            provenance=EdgeProvenance(causes=frozenset({DependencyCause.direct_ref})),
        )
    index = build_taco_index(graph)
    chains = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr_chain]
    assert len(chains) == 1
    assert_taco_parity(graph, index)


def test_rr_chain_workbook_parity(tmp_path: Path) -> None:
    path = tmp_path / "chain.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    ws.write_number(2, 15, 1.0)
    for row in range(4, 10):
        ws.write_formula(row - 1, 15, f"=P{row - 1}+1")
    wb.close()

    graph = create_dependency_graph(path, ["Data!P4:P9"], load_values=False)
    index = build_taco_index(graph)
    assert any(e.meta.kind == PatternKind.rr_chain for e in index.compressed_edges)
    assert_taco_parity(graph, index)
