"""Tests for TACO RF, FR, and FF range-pattern compression along rows."""

from __future__ import annotations

from pathlib import Path

import fastpyxl.utils.cell
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import (
    PatternKind,
    RangeRef,
    build_taco_index,
    materialize_precedents,
)

from .parity_helpers import assert_taco_parity


def test_rf_running_sum_row_parity(tmp_path: Path) -> None:
    path = tmp_path / "rf_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    row = 9
    tail_col = 16  # P
    first_col, last_col = 6, 12  # F through L
    for col_i in range(5, tail_col):
        ws.write_number(row - 1, col_i - 1, 10.0)
    tail = fastpyxl.utils.cell.get_column_letter(tail_col)
    for col_i in range(first_col, last_col + 1):
        start_col = fastpyxl.utils.cell.get_column_letter(col_i - 1)
        ws.write_formula(row - 1, col_i - 1, f"=SUM({start_col}{row}:${tail}${row})")
    wb.close()

    first = fastpyxl.utils.cell.get_column_letter(first_col)
    last = fastpyxl.utils.cell.get_column_letter(last_col)
    graph = create_dependency_graph(path, [f"Data!{first}{row}:{last}{row}"], load_values=False)
    index = build_taco_index(graph)
    rf = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rf]
    assert len(rf) == 1
    assert rf[0].dependent == RangeRef.row_span("Data", row, first, last)
    dep_col = fastpyxl.utils.cell.get_column_letter(8)  # H
    assert materialize_precedents(index, f"Data!{dep_col}{row}") == {
        f"Data!{fastpyxl.utils.cell.get_column_letter(c)}{row}" for c in range(7, tail_col + 1)
    }
    assert_taco_parity(graph, index)


def test_fr_ytd_row_parity(tmp_path: Path) -> None:
    path = tmp_path / "fr_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    row = 9
    head_col = 5  # E
    first_col, last_col = 6, 12  # F through L
    head = fastpyxl.utils.cell.get_column_letter(head_col)
    for col_i in range(head_col, last_col + 1):
        ws.write_number(row - 1, col_i - 1, float(col_i))
    for col_i in range(first_col, last_col + 1):
        dep_col = fastpyxl.utils.cell.get_column_letter(col_i)
        ws.write_formula(row - 1, col_i - 1, f"=SUM(${head}${row}:{dep_col}{row})")
    wb.close()

    first = fastpyxl.utils.cell.get_column_letter(first_col)
    last = fastpyxl.utils.cell.get_column_letter(last_col)
    graph = create_dependency_graph(path, [f"Data!{first}{row}:{last}{row}"], load_values=False)
    index = build_taco_index(graph)
    fr = [e for e in index.compressed_edges if e.meta.kind == PatternKind.fr]
    assert len(fr) == 1
    assert fr[0].dependent == RangeRef.row_span("Data", row, first, last)
    assert materialize_precedents(index, "Data!H9") == {
        f"Data!{fastpyxl.utils.cell.get_column_letter(c)}9" for c in range(head_col, 9)
    }
    assert_taco_parity(graph, index)


def test_ff_vlookup_table_row_parity(tmp_path: Path) -> None:
    path = tmp_path / "ff_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")
    row = 9
    keys = ["a", "b", "c", "d", "e"]
    table_first_col, table_last_col = 13, 14  # M, N
    key_first_col = 10  # J
    for i, key in enumerate(keys):
        key_col_i = key_first_col + i
        ws.write_string(row - 1, key_col_i - 1, key)
        ws.write_string(row + i - 1, table_first_col - 1, key)
        ws.write_number(row + i - 1, table_last_col - 1, float(i + 1))
    for i in range(len(keys)):
        key_col = fastpyxl.utils.cell.get_column_letter(key_first_col + i)
        ws.write_formula(
            row - 1,
            key_first_col + i,
            f"=VLOOKUP({key_col}{row},"
            f"${fastpyxl.utils.cell.get_column_letter(table_first_col)}${row}:"
            f"${fastpyxl.utils.cell.get_column_letter(table_last_col)}${row + len(keys) - 1},"
            f"2,FALSE)",
        )
    wb.close()

    first = fastpyxl.utils.cell.get_column_letter(key_first_col + 1)
    last = fastpyxl.utils.cell.get_column_letter(key_first_col + len(keys))
    graph = create_dependency_graph(path, [f"Data!{first}{row}:{last}{row}"], load_values=False)
    index = build_taco_index(graph)
    ff = [e for e in index.compressed_edges if e.meta.kind == PatternKind.ff]
    assert len(ff) == 1
    table = {
        f"Data!{fastpyxl.utils.cell.get_column_letter(table_first_col)}{r}"
        for r in range(row, row + len(keys))
    } | {
        f"Data!{fastpyxl.utils.cell.get_column_letter(table_last_col)}{r}"
        for r in range(row, row + len(keys))
    }
    assert table <= materialize_precedents(index, "Data!M9")
    assert_taco_parity(graph, index)
