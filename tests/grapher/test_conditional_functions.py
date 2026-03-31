from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph


def test_choose_branches_are_guarded(tmp_path: Path) -> None:
    excel_path = tmp_path / "choose.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")

    ws.write_number(0, 2, 1)  # C1
    ws.write_number(0, 1, 10)  # B1
    ws.write_number(0, 3, 20)  # D1
    ws.write_formula(0, 0, "=CHOOSE($C$1,B1,D1)", None, 10)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    deps = graph.dependencies("Sheet1!A1")
    # Selector cell (C1) is unconditional; branch values are guarded.
    assert "Sheet1!C1" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!D1" in deps
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!C1").get("guard") is None
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!B1").get("guard") is not None
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!D1").get("guard") is not None


def test_switch_branches_are_guarded(tmp_path: Path) -> None:
    import fastpyxl

    excel_path = tmp_path / "switch.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["C1"].value = 2
    ws["B1"].value = 10
    ws["D1"].value = 20
    ws["E1"].value = 30
    ws["A1"].value = "=SWITCH($C$1,1,B1,2,D1,E1)"
    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    deps = graph.dependencies("Sheet1!A1")
    assert "Sheet1!C1" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!D1" in deps
    assert "Sheet1!E1" in deps
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!C1").get("guard") is None
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!B1").get("guard") is not None
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!D1").get("guard") is not None
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!E1").get("guard") is not None


def test_ifs_branches_are_guarded(tmp_path: Path) -> None:
    excel_path = tmp_path / "ifs.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")

    ws.write_number(0, 2, 0)  # C1
    ws.write_number(0, 1, 10)  # B1
    ws.write_number(0, 3, 20)  # D1
    # If C1=0 -> B1, else if C1=1 -> D1, else -> 0
    ws.write_formula(0, 0, "=IFS($C$1=0,B1,$C$1=1,D1,TRUE,0)", None, 10)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    deps = graph.dependencies("Sheet1!A1")
    assert "Sheet1!C1" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!D1" in deps
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!C1").get("guard") is None
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!B1").get("guard") is not None
    assert graph.edge_attrs("Sheet1!A1", "Sheet1!D1").get("guard") is not None
