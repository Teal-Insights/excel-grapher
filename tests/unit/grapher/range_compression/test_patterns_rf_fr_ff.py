"""Tests for TACO RF, FR, and FF range-pattern compression."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import (
    PatternKind,
    build_taco_index,
    materialize_precedents,
)

from .parity_helpers import assert_taco_parity


def test_rf_running_sum_column_parity(tmp_path: Path) -> None:
    path = tmp_path / "rf_column.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    tail = 20
    for row in range(3, tail + 1):
        ws.write_number(row - 1, 4, 5.0)
    for row in range(3, 13):
        ws.write_formula(row - 1, 5, f"=SUM(E{row}:$E${tail})")
    wb.close()

    graph = create_dependency_graph(
        path, ["Data!F3:F12"], load_values=False, store_raw_formula=True
    )
    index = build_taco_index(graph)
    rf = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rf]
    assert len(rf) == 1
    assert materialize_precedents(index, "Data!F5") == {f"Data!E{r}" for r in range(5, tail + 1)}
    assert_taco_parity(graph, index)


def test_fr_ytd_column_parity(tmp_path: Path) -> None:
    path = tmp_path / "fr_column.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    first, last = 3, 12
    for row in range(first, last + 1):
        ws.write_number(row - 1, 6, float(row))
        ws.write_formula(row - 1, 7, f"=SUM($G${first}:G{row})")
    wb.close()

    graph = create_dependency_graph(
        path, [f"Data!H{first}:H{last}"], load_values=False, store_raw_formula=True
    )
    index = build_taco_index(graph)
    fr = [e for e in index.compressed_edges if e.meta.kind == PatternKind.fr]
    assert len(fr) == 1
    assert materialize_precedents(index, "Data!H5") == {f"Data!G{r}" for r in range(first, 6)}
    assert_taco_parity(graph, index)


def test_ff_vlookup_table_parity(tmp_path: Path) -> None:
    path = tmp_path / "ff_column.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    keys = ["a", "b", "c"]
    for i, key in enumerate(keys):
        row = 3 + i
        ws.write_string(row - 1, 9, key)
        ws.write_string(row - 1, 12, key)
        ws.write_number(row - 1, 13, float(i + 1))
    for row in range(3, 8):
        ws.write_formula(
            row - 1,
            10,
            f"=VLOOKUP(J{row},$M$3:$N$5,2,FALSE)",
        )
    wb.close()

    graph = create_dependency_graph(path, ["Data!K3:K7"], load_values=False, store_raw_formula=True)
    index = build_taco_index(graph)
    ff = [e for e in index.compressed_edges if e.meta.kind == PatternKind.ff]
    assert len(ff) == 1
    table = {f"Data!M{r}" for r in range(3, 6)} | {f"Data!N{r}" for r in range(3, 6)}
    assert table <= materialize_precedents(index, "Data!K4")
    assert_taco_parity(graph, index)
