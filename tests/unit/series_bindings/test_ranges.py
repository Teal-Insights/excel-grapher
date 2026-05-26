"""Unit tests for data_range expansion (targets + named ranges)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.resolver import build_named_range_map
from excel_grapher.series_bindings import expand_data_range, expand_data_range_for_graph


def test_expand_local_range_inputs_f5_j5() -> None:
    addresses = expand_data_range("Inputs!F5:J5")
    assert addresses == ["Inputs!F5", "Inputs!G5", "Inputs!H5", "Inputs!I5", "Inputs!J5"]


def test_expand_both_end_sheet_qualified_form() -> None:
    addresses = expand_data_range("Sheet1!C2:Sheet1!E2")
    assert addresses == ["Sheet1!C2", "Sheet1!D2", "Sheet1!E2"]


def test_expand_quoted_sheet_both_ends(tmp_path: Path) -> None:
    path = tmp_path / "quoted.xlsx"
    wb = xlsxwriter.Workbook(path)
    wb.add_worksheet("My Sheet")
    wb.close()
    addresses = expand_data_range("'My Sheet'!A1:'My Sheet'!B2", workbook=path)
    assert sorted(addresses) == sorted(
        ["'My Sheet'!A1", "'My Sheet'!A2", "'My Sheet'!B1", "'My Sheet'!B2"]
    )


def test_expand_defined_name_range(tmp_path: Path) -> None:
    path = tmp_path / "named.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write_number("F5", 1)
    ws.write_number("G5", 2)
    wb.define_name("PrimaryRow", "Inputs!$F$5:$G$5")
    wb.close()

    maps = build_named_range_map(fastpyxl.load_workbook(path, data_only=False, read_only=True))
    addresses = expand_data_range(
        "PrimaryRow",
        sheetnames=["Inputs"],
        named_ranges=maps.cell_map,
        named_range_ranges=maps.range_map,
    )
    assert addresses == ["Inputs!F5", "Inputs!G5"]


def test_expand_defined_name_requires_maps_or_workbook() -> None:
    with pytest.raises(ValueError, match="Unknown defined name"):
        expand_data_range("NoSuchName")


def test_expand_data_range_for_graph_uses_graph_named_ranges(tmp_path: Path) -> None:
    path = tmp_path / "named.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write_number("F5", 1)
    wb.define_name("CellOne", "Inputs!$F$5")
    wb.close()

    graph = create_dependency_graph(path, ["Inputs!F5"], load_values=True)
    assert expand_data_range_for_graph(graph, "CellOne") == ["Inputs!F5"]
