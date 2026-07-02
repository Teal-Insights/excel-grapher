"""Tests for TACO RR-Chain pattern compression along rows (fill-right)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl.utils.cell
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.range_compression import PatternKind, RangeRef, build_taco_index

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


def test_rr_chain_row_fill_right_manual_graph() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("Sheet1!E9", formula=None, is_leaf=True))
    dr = DependencyCause.direct_ref
    for col in ("F", "G", "H", "I", "J"):
        prev = chr(ord(col) - 1)
        graph.add_node(_make_node(f"Sheet1!{col}9", formula=f"={prev}9+1"))
        graph.add_edge(
            f"Sheet1!{col}9",
            f"Sheet1!{prev}9",
            provenance=EdgeProvenance(causes=frozenset({dr})),
        )

    index = build_taco_index(graph)
    chains = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr_chain]
    assert len(chains) == 1
    edge = chains[0]
    assert edge.dependent == RangeRef.row_span("Sheet1", 9, "F", "J")
    assert edge.precedent == RangeRef.row_span("Sheet1", 9, "E", "I")
    assert_taco_parity(graph, index)


def test_rr_chain_row_workbook_parity(tmp_path: Path) -> None:
    path = tmp_path / "chain_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    row = 9
    ws.write_number(row - 1, 4, 1.0)  # E9
    for col_i in range(6, 11):  # F through J
        prev = fastpyxl.utils.cell.get_column_letter(col_i - 1)
        ws.write_formula(row - 1, col_i - 1, f"={prev}{row}+1")
    wb.close()

    graph = create_dependency_graph(path, ["Data!F9:J9"], load_values=False)
    index = build_taco_index(graph)
    assert any(e.meta.kind == PatternKind.rr_chain for e in index.compressed_edges)
    assert_taco_parity(graph, index)
