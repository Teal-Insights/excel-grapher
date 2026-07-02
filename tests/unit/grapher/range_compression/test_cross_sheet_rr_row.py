"""Cross-sheet TACO RR pattern compression along rows."""

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


def test_cross_sheet_rr_row_parity(tmp_path: Path) -> None:
    path = tmp_path / "cross_rr_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    data = wb.add_worksheet("Data")
    report = wb.add_worksheet("Report")
    row = 9
    for col_i in range(6, 11):
        src = get_column_letter(col_i)
        data.write_number(row - 1, col_i - 1, float(col_i))
        report.write_formula(row - 1, col_i - 1, f"=Data!{src}{row}")
    wb.close()

    graph = create_dependency_graph(path, ["Report!F9:J9"], load_values=False)
    index = build_taco_index(graph)
    rr = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr) == 1
    assert rr[0].dependent == RangeRef.row_span("Report", row, "F", "J")
    assert rr[0].precedent.sheet == "Data"
    assert rr[0].dependent.sheet == "Report"
    assert materialize_precedents(index, "Report!H9") == {"Data!H9"}
    assert_taco_parity(graph, index)
