"""CHOOSE/IFS-style formulas yield guarded dependency edges in built graphs (integration).

Uses small on-disk workbooks to assert selector cells versus guarded branch edges
match how analysts read conditional spreadsheet dependencies.
"""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig


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
    deps = graph.get_dependencies("Sheet1!A1")
    # Selector cell (C1) is unconditional; branch values are guarded.
    assert "Sheet1!C1" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!D1" in deps
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!C1") is None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!B1") is not None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!D1") is not None


def test_choose_literal_index_drops_unselected_branch(tmp_path: Path) -> None:
    excel_path = tmp_path / "choose_literal.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1
    ws.write_number(0, 3, 20)  # D1
    ws.write_formula(0, 0, "=CHOOSE(1,B1,D1)", None, 10)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    assert set(graph.get_dependencies("Sheet1!A1")) == {"Sheet1!B1"}
    graph.validate_consistency()


def test_choose_domain_keeps_feasible_refs_and_drops_ref_error(tmp_path: Path) -> None:
    """A constrained index keeps only the alternatives it can select."""
    excel_path = tmp_path / "choose_domain.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 10)  # A1
    ws.write_number(0, 1, 20)  # B1
    ws.write_number(0, 3, 40)  # D1
    ws.write_number(0, 4, 1)  # E1
    ws.write_formula(0, 2, "=IF(E1=0,0,CHOOSE(E1,A1,B1,#REF!,D1))", None, 10)
    wb.close()

    config = DynamicRefConfig.from_constraints({"Sheet1!E1": Literal[1, 2]})
    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!C1"],
        load_values=False,
        dynamic_refs=config,
        capture_dependency_provenance=True,
    )
    assert set(graph.get_dependencies("Sheet1!C1")) == {
        "Sheet1!E1",
        "Sheet1!A1",
        "Sheet1!B1",
    }
    assert graph.get_edge_guard("Sheet1!C1", "Sheet1!E1") is None
    assert graph.get_edge_guard("Sheet1!C1", "Sheet1!A1") is not None
    assert graph.get_edge_guard("Sheet1!C1", "Sheet1!B1") is not None
    graph.validate_consistency()


def test_choose_row_index_uses_the_host_row(tmp_path: Path) -> None:
    excel_path = tmp_path / "choose_row.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1
    ws.write_number(1, 1, 20)  # B2
    # ROW() on A2 is 2, so only the second alternative is read.
    ws.write_formula(1, 0, "=CHOOSE(ROW(),B1,B2)", None, 20)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A2"], load_values=False)
    assert set(graph.get_dependencies("Sheet1!A2")) == {"Sheet1!B2"}
    graph.validate_consistency()


def test_choose_without_index_domain_keeps_every_alternative(tmp_path: Path) -> None:
    excel_path = tmp_path / "choose_open.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 10)  # A1
    ws.write_number(0, 1, 20)  # B1
    ws.write_number(0, 3, 40)  # D1
    ws.write_number(0, 4, 1)  # E1
    ws.write_formula(0, 2, "=CHOOSE(E1,A1,B1,#REF!,D1)", None, 10)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!C1"],
        load_values=False,
        dynamic_refs=DynamicRefConfig.from_constraints({}),
    )
    assert set(graph.get_dependencies("Sheet1!C1")) == {
        "Sheet1!E1",
        "Sheet1!A1",
        "Sheet1!B1",
        "Sheet1!D1",
    }
    graph.validate_consistency()


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
    deps = graph.get_dependencies("Sheet1!A1")
    assert "Sheet1!C1" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!D1" in deps
    assert "Sheet1!E1" in deps
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!C1") is None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!B1") is not None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!D1") is not None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!E1") is not None


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
    deps = graph.get_dependencies("Sheet1!A1")
    assert "Sheet1!C1" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!D1" in deps
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!C1") is None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!B1") is not None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!D1") is not None
