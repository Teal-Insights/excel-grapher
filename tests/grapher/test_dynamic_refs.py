"""Regression tests for dynamic reference parsing (OFFSET/INDIRECT)."""
from __future__ import annotations

import warnings
from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph


def _build_offset_named_range_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(0, 3, 20)  # D1 target
    ws.write_formula(0, 2, "=1+1", None, 2)  # C1 (LANG) cached value = 2
    ws.write_formula(0, 0, "=OFFSET(B1,0,LANG)+OFFSET(B1,0,LANG)", None, 40)
    wb.define_name("LANG", "Sheet1!$C$1")
    wb.close()


def _build_offset_index_row_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    lookup = wb.add_worksheet("lookup")
    sheet = wb.add_worksheet("Sheet1")

    # Named range Country_list -> lookup!C4:C6
    wb.define_name("Country_list", "lookup!$C$4:$C$6")

    # Seed lookup range and the shifted target column
    lookup.write_number(3, 1, 111)  # B4
    lookup.write_number(3, 2, 222)  # C4

    # In B2, ROW()-ROW($B$2)+1 resolves to 1
    sheet.write_formula(
        1,
        1,
        "=OFFSET(INDEX(Country_list,ROW()-ROW($B$2)+1,1),0,-1)",
        None,
        111,
    )
    wb.close()


def _build_offset_with_arg_ref_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    start = wb.add_worksheet("START")

    ws.write_number(0, 1, 5)  # B1 base
    ws.write_number(0, 2, 99)  # C1 target
    start.write_number(9, 12, 1)  # M10 -> column offset of 1

    ws.write_formula(0, 0, "=OFFSET(B1,0,START!M10)", None, 99)
    wb.close()


def test_offset_with_cached_named_range_warns_once(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_named_range.xlsx"
    _build_offset_named_range_workbook(excel_path)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)

    deps = graph.dependencies("Sheet1!A1")
    # A1 = OFFSET(B1,0,LANG)+OFFSET(B1,0,LANG); LANG = Sheet1!C1. Deps include C1 (offset arg) and D1 (resolved target).
    assert deps == {"Sheet1!C1", "Sheet1!D1"}

    cache_warnings = [
        w for w in caught if "cached workbook values" in str(w.message)
    ]
    assert len(cache_warnings) == 1


def test_offset_index_row_resolves_named_range(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_index_row.xlsx"
    _build_offset_index_row_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!B2"], load_values=False)
    deps = graph.dependencies("Sheet1!B2")
    assert deps == {"lookup!B4"}


def test_offset_argument_references_are_dependencies(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_arg_ref.xlsx"
    _build_offset_with_arg_ref_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)
    deps = graph.dependencies("Sheet1!A1")
    assert deps == {"Sheet1!C1", "START!M10"}
