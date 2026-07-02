"""Cross-sheet RF, FR, and FF TACO pattern tests along rows."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter
from fastpyxl.utils.cell import get_column_letter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import (
    PatternKind,
    RangeRef,
    build_taco_index,
    materialize_precedents,
)

from .parity_helpers import assert_taco_parity


def test_cross_sheet_rf_fr_ff_row_workbook(tmp_path: Path) -> None:
    path = tmp_path / "cross_patterns_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    data = wb.add_worksheet("Data")
    report = wb.add_worksheet("Report")
    row = 9
    span = 5
    rf_data_first = 12
    rf_tail_col = 22
    rf_form_first = 23
    fr_head_col = 29
    ff_key_first = 52
    ff_table_first = 55
    keys = ["a", "b", "c", "d", "e"]

    for col_i in range(rf_data_first, rf_tail_col + 1):
        data.write_number(row - 1, col_i - 1, 10.0)
    for col_i in range(fr_head_col, fr_head_col + span + 1):
        data.write_number(row - 1, col_i - 1, float(col_i))
    for i, key in enumerate(keys):
        data.write_string(row - 1, ff_key_first + i - 1, key)
        data.write_string(row + i - 1, ff_table_first - 1, key)
        data.write_number(row + i - 1, ff_table_first, float(i + 1))

    for offset, col_i in enumerate(range(rf_form_first, rf_form_first + span)):
        start_col = get_column_letter(rf_data_first + offset)
        tail_col = get_column_letter(rf_tail_col)
        report.write_formula(
            row - 1,
            col_i - 1,
            f"=SUM(Data!{start_col}{row}:Data!${tail_col}${row})",
        )
    for _offset, col_i in enumerate(range(fr_head_col + span + 1, fr_head_col + 2 * span + 1)):
        head = get_column_letter(fr_head_col)
        tail = get_column_letter(col_i)
        report.write_formula(
            row - 1,
            col_i - 1,
            f"=SUM(Data!${head}${row}:Data!{tail}{row})",
        )
    for i in range(span):
        key_col = get_column_letter(ff_key_first + i)
        report.write_formula(
            row - 1,
            ff_key_first + i,
            f"=VLOOKUP({key_col}{row},"
            f"Data!${get_column_letter(ff_table_first)}${row}:"
            f"Data!${get_column_letter(ff_table_first + 1)}${row + len(keys) - 1},"
            f"2,FALSE)",
        )
    wb.close()

    rf_last = get_column_letter(rf_form_first + span - 1)
    fr_first = get_column_letter(fr_head_col + span + 1)
    fr_last = get_column_letter(fr_head_col + 2 * span)
    ff_first = get_column_letter(ff_key_first + 1)
    ff_last = get_column_letter(ff_key_first + span)
    graph = create_dependency_graph(
        path,
        [
            f"Report!{get_column_letter(rf_form_first)}{row}:{rf_last}{row}",
            f"Report!{fr_first}{row}:{fr_last}{row}",
            f"Report!{ff_first}{row}:{ff_last}{row}",
        ],
        load_values=False,
    )
    index = build_taco_index(graph)
    kinds = {e.meta.kind for e in index.compressed_edges}
    assert PatternKind.rf in kinds
    assert PatternKind.fr in kinds
    assert PatternKind.ff in kinds
    rf = next(e for e in index.compressed_edges if e.meta.kind == PatternKind.rf)
    assert rf.precedent.sheet == "Data"
    assert rf.dependent.sheet == "Report"
    assert rf.dependent == RangeRef.row_span(
        "Report", row, get_column_letter(rf_form_first), rf_last
    )
    assert {k.split("!")[0] for k in materialize_precedents(index, f"Report!{rf_last}{row}")} == {
        "Data"
    }
    assert_taco_parity(graph, index)
