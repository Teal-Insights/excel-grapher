"""Cross-sheet RF, FR, and FF TACO pattern tests."""

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


def test_cross_sheet_rf_fr_ff_workbook(tmp_path: Path) -> None:
    path = tmp_path / "cross_patterns.xlsx"
    wb = xlsxwriter.Workbook(path)
    data = wb.add_worksheet("Data")
    report = wb.add_worksheet("Report")
    tail = 11
    first, last = 3, 7
    keys = ["a", "b", "c", "d", "e"]
    for er in range(first, tail + 1):
        data.write_number(er - 1, 4, 10.0)
    for er in range(first, last + 1):
        data.write_number(er - 1, 6, float(er))
        data.write_string(er - 1, 12, keys[er - first])
        data.write_number(er - 1, 13, float(er))
        report.write_formula(er - 1, 5, f"=SUM(Data!E{er}:Data!$E${tail})")
        report.write_formula(er - 1, 7, f"=SUM(Data!$G${first}:Data!G{er})")
        report.write_string(er - 1, 9, keys[er - first])
        report.write_formula(
            er - 1,
            10,
            f"=VLOOKUP(J{er},Data!$M${first}:Data!$N${first + len(keys) - 1},2,FALSE)",
        )
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Report!F3:F7", "Report!H3:H7", "Report!K3:K7"],
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
    assert {k.split("!")[0] for k in materialize_precedents(index, "Report!F5")} == {"Data"}
    assert_taco_parity(graph, index)
