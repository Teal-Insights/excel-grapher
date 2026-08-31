"""Defined-name table write-back; formula tokens stay expanded A1 (#569)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest
from fastpyxl.workbook.defined_name import DefinedName

from excel_grapher.core.formula_ast import FormulaStyle, parse_preserving_axes, render_formula
from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.grapher.resolver import build_named_range_map


def _write_tax_rate_source(path: Path) -> None:
    wb = fastpyxl.Workbook()
    settings = wb.active
    assert settings is not None
    settings.title = "Settings"
    settings["B2"] = 0.2
    sheet1 = wb.create_sheet("Sheet1")
    sheet1["A1"] = 10
    sheet1["C1"] = "=A1*TaxRate"
    wb.defined_names.add(DefinedName(name="TaxRate", attr_text="Settings!$B$2"))
    wb.save(path)
    wb.close()


def _defined_names(path: Path) -> dict[str, str]:
    wb = fastpyxl.load_workbook(path)
    try:
        return {
            name: defn.attr_text
            for name, defn in wb.defined_names.items()
            if isinstance(defn.attr_text, str)
        }
    finally:
        wb.close()


def _named_range_maps(path: Path):
    wb = fastpyxl.load_workbook(path, data_only=False, read_only=True)
    try:
        return build_named_range_map(wb)
    finally:
        wb.close()


def test_write_workbook_keeps_defined_name_table_and_expands_formula_tokens(
    tmp_path: Path,
) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_tax_rate_source(source)

    graph = create_dependency_graph(source, ["Sheet1!C1"], load_values=True, store_raw_formula=True)
    node = graph.get_node("Sheet1!C1")
    assert node is not None and node.formula_ast is not None
    assert node.formula == "=A1*TaxRate"
    assert graph.named_ranges is not None
    assert graph.named_ranges["TaxRate"] == ("Settings", "B2")

    expected = render_formula(node.formula_ast, anchor=node.address, style=FormulaStyle.A1_EXCEL)
    assert "TaxRate" not in expected

    write_workbook(graph, dest)

    maps = _named_range_maps(dest)
    assert maps.cell_map["TaxRate"] == ("Settings", "B2")
    assert _defined_names(dest)["TaxRate"] == "Settings!$B$2"

    wb = fastpyxl.load_workbook(dest)
    try:
        assert wb["Sheet1"]["C1"].value == expected
        assert wb["Sheet1"]["C1"].value != "=A1*TaxRate"
    finally:
        wb.close()

    restored = create_dependency_graph(dest, ["Sheet1!C1"], store_raw_formula=True)
    after = restored.get_node("Sheet1!C1")
    assert after is not None and after.formula_ast is not None
    assert after.formula == expected
    assert restored.named_ranges is not None
    assert restored.named_ranges["TaxRate"] == ("Settings", "B2")


def test_write_workbook_writes_named_range_rectangles(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 1
    ws["B1"] = 2
    ws["C1"] = "=SUM(Inputs)"
    wb.defined_names.add(DefinedName(name="Inputs", attr_text="Sheet1!$A$1:$B$1"))
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(source, ["Sheet1!C1"])
    assert graph.named_range_ranges is not None
    assert graph.named_range_ranges["Inputs"] == ("Sheet1", "A1", "B1")
    write_workbook(graph, dest)

    maps = _named_range_maps(dest)
    assert maps.range_map["Inputs"] == ("Sheet1", "A1", "B1")
    assert _defined_names(dest)["Inputs"] == "Sheet1!$A$1:$B$1"
    node = graph.get_node("Sheet1!C1")
    assert node is not None and node.formula_ast is not None
    wb_out = fastpyxl.load_workbook(dest)
    try:
        assert wb_out["Sheet1"]["C1"].value == render_formula(
            node.formula_ast, anchor=node.address, style=FormulaStyle.A1_EXCEL
        )
    finally:
        wb_out.close()


def test_write_workbook_quotes_sheet_names_in_defined_names(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    rates = wb.active
    assert rates is not None
    rates.title = "Rate Sheet"
    rates["B2"] = 0.1
    sheet1 = wb.create_sheet("Sheet1")
    sheet1["A1"] = 5
    sheet1["C1"] = "=A1*TaxRate"
    wb.defined_names.add(DefinedName(name="TaxRate", attr_text="'Rate Sheet'!$B$2"))
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(source, ["Sheet1!C1"])
    write_workbook(graph, dest)
    maps = _named_range_maps(dest)
    assert maps.cell_map["TaxRate"] == ("Rate Sheet", "B2")
    assert _defined_names(dest)["TaxRate"] == "'Rate Sheet'!$B$2"


def test_write_workbook_include_defined_names_false_omits_name_table(
    tmp_path: Path,
) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_tax_rate_source(source)
    graph = create_dependency_graph(source, ["Sheet1!C1"])
    write_workbook(graph, dest, include_defined_names=False)
    assert _defined_names(dest) == {}
    node = graph.get_node("Sheet1!C1")
    assert node is not None and node.formula_ast is not None
    wb_out = fastpyxl.load_workbook(dest)
    try:
        assert wb_out["Sheet1"]["C1"].value == render_formula(
            node.formula_ast, anchor=node.address, style=FormulaStyle.A1_EXCEL
        )
    finally:
        wb_out.close()


def test_write_workbook_does_not_invent_names_from_expanded_ast(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Settings", "Sheet1"])
    graph.add_node(make_cell_node("Settings", "B", 2, value=0.2, is_leaf=True))
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=10, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "C",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=A1*Settings!$B$2", anchor="Sheet1!C1"),
        )
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    assert _defined_names(dest) == {}


def test_write_workbook_refuses_name_that_is_both_cell_and_range(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True))
    graph.named_ranges = {"TaxRate": ("Sheet1", "A1")}
    graph.named_range_ranges = {"TaxRate": ("Sheet1", "A1", "B1")}
    with pytest.raises(ValueError, match="TaxRate"):
        write_workbook(graph, tmp_path / "overlap.xlsx")
    assert not (tmp_path / "overlap.xlsx").exists()


def test_write_workbook_refuses_defined_name_that_is_not_a_cell_or_range(
    tmp_path: Path,
) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True))
    graph.named_ranges = {"TaxRate": ("Sheet1", "1+1")}
    with pytest.raises(ValueError, match="TaxRate"):
        write_workbook(graph, tmp_path / "formula-name.xlsx")
    assert not (tmp_path / "formula-name.xlsx").exists()


def test_write_workbook_refuses_empty_defined_name(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True))
    graph.named_ranges = {"": ("Sheet1", "A1")}
    with pytest.raises(ValueError, match="defined name"):
        write_workbook(graph, tmp_path / "empty-name.xlsx")
    assert not (tmp_path / "empty-name.xlsx").exists()


def test_write_workbook_projection_writes_original_name_table(tmp_path: Path) -> None:
    from excel_grapher.exporter import IdentityTransitCompression

    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    settings = wb.active
    assert settings is not None
    settings.title = "Settings"
    settings["B2"] = 0.2
    outputs = wb.create_sheet("Outputs")
    outputs["B12"] = "=TaxRate"
    outputs["B14"] = "=Outputs!B12+1"
    wb.defined_names.add(DefinedName(name="TaxRate", attr_text="Settings!$B$2"))
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(
        source,
        ["Outputs!B14"],
        capture_dependency_provenance=True,
        load_values=True,
    )
    projection = IdentityTransitCompression().project(graph)
    write_workbook(projection, dest)
    maps = _named_range_maps(dest)
    assert maps.cell_map["TaxRate"] == ("Settings", "B2")
