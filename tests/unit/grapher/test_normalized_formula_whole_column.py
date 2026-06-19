from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher import create_dependency_graph


def test_normalized_formula_preserves_whole_column_shorthand(tmp_path: Path) -> None:
    excel_path = tmp_path / "norm_whole_col.xlsx"
    wb = fastpyxl.Workbook()
    data = wb.create_sheet("Data")
    sheet1 = wb.active
    sheet1.title = "Sheet1"
    sheet1["B1"].value = '=MATCH("x",Data!A:A,0)'
    data["A1"].value = "x"
    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!B1"], load_values=False)
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.normalized_formula == '=MATCH("x",Data!A:A,0)'
