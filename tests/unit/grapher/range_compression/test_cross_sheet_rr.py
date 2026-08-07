"""Cross-sheet TACO RR pattern compression tests."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import (
    PatternKind,
    RangeRef,
    build_taco_index,
    materialize_precedents,
)

from .parity_helpers import assert_taco_parity


def test_cross_sheet_rr_column_parity(tmp_path: Path) -> None:
    path = tmp_path / "cross_rr.xlsx"
    wb = xlsxwriter.Workbook(path)
    data = wb.add_worksheet("Data")
    report = wb.add_worksheet("Report")
    for row in range(3, 8):
        data.write_number(row - 1, 1, float(row))
        report.write_formula(row - 1, 3, f"=Data!B{row}")
    wb.close()

    graph = create_dependency_graph(
        path, ["Report!D3:D7"], load_values=False, store_raw_formula=True
    )
    index = build_taco_index(graph)
    rr = [e for e in index.compressed_edges if e.meta.kind == PatternKind.rr]
    assert len(rr) == 1
    assert rr[0].dependent.sheet == "Report"
    assert rr[0].precedent.sheet == "Data"
    assert materialize_precedents(index, "Report!D5") == {"Data!B5"}
    assert_taco_parity(graph, index)


def test_cross_sheet_rr_materialize_uses_precedent_sheet() -> None:
    from excel_grapher.grapher.range_compression.patterns import rr_materialize_precedent

    dep = RangeRef(sheet="Report", min_col="D", min_row=3, max_col="D", max_row=7)
    prec = RangeRef(sheet="Data", min_col="B", min_row=3, max_col="B", max_row=7)
    assert rr_materialize_precedent(dep, prec, "Report!D5") == "Data!B5"
