from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph


def _write_match_column_workbook(path: Path, *, interior_row: int, used_rows: int) -> None:
    wb = xlsxwriter.Workbook(path)
    data = wb.add_worksheet("Data")
    sheet1 = wb.add_worksheet("Sheet1")
    data.write_string(interior_row - 1, 0, "SEA")
    sheet1.write_formula(0, 1, '=MATCH("SEA",Data!A:A,0)', None, 0)
    wb.close()


def test_whole_column_deps_include_interior_cell_not_corners_only(tmp_path: Path) -> None:
    excel_path = tmp_path / "match_interior.xlsx"
    _write_match_column_workbook(excel_path, interior_row=25, used_rows=50)

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!B1"],
        load_values=True,
        max_range_cells=2,
    )
    deps = graph.get_dependencies("Sheet1!B1")
    assert "Data!A25" in deps
    assert len([d for d in deps if d.startswith("Data!A")]) >= 3


def test_whole_column_deps_bounded_to_used_range_not_excel_max(tmp_path: Path) -> None:
    excel_path = tmp_path / "match_bounded.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    data = wb.add_worksheet("Data")
    sheet1 = wb.add_worksheet("Sheet1")
    for row in range(50):
        data.write_string(row, 0, "x")
    sheet1.write_formula(0, 1, '=MATCH("x",Data!A:A,0)', None, 1)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!B1"],
        load_values=True,
        max_range_cells=2,
    )
    deps = graph.get_dependencies("Sheet1!B1")
    assert "Data!A50" in deps
    assert "Data!A1048576" not in deps


def test_rectangular_range_still_uses_corner_fallback_issue_56(tmp_path: Path) -> None:
    excel_path = tmp_path / "rect_corner_cap.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    ws.write_number(0, 2, 1)
    ws.write_formula(0, 3, "=SUM(Sheet1!A1:C1)", None, 2)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!D1"],
        load_values=False,
        max_range_cells=2,
    )
    deps = graph.get_dependencies("Sheet1!D1")
    assert "Sheet1!A1" in deps
    assert "Sheet1!C1" in deps
    assert "Sheet1!B1" not in deps
