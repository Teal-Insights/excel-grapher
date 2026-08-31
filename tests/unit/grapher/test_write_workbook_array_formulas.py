"""Array-formula provenance and CSE/dynamic write-back (#565)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest
import xlsxwriter
from fastpyxl.worksheet.formula import ArrayFormula

from excel_grapher.core.formula_ast import FormulaStyle, parse_preserving_axes, render_formula
from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.grapher.cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    dependency_graph_from_json,
    dependency_graph_to_json,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import copy_node, make_cell_node


def _write_cse(path: Path, formula: str, *, ref: str = "E1:E1") -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(3):
        ws.write_number(row, 0, row + 1)  # A1:A3
        ws.write_number(row, 1, 10 * (row + 1))  # B1:B3
    ws.write_array_formula(ref, formula, None, 0)
    wb.close()
    return path


def _write_dynamic(path: Path, formula: str, *, ref: str = "E1:E3") -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(3):
        ws.write_number(row, 0, row + 1)
        ws.write_number(row, 1, 10 * (row + 1))
    ws.write_dynamic_array_formula(ref, formula, None, 0)
    wb.close()
    return path


def _array_cell(path: Path, coord: str = "E1") -> ArrayFormula:
    wb = fastpyxl.load_workbook(path)
    try:
        raw = wb["Sheet1"][coord].value
    finally:
        wb.close()
    assert isinstance(raw, ArrayFormula)
    return raw


def test_extract_stores_cse_array_formula_provenance(tmp_path: Path) -> None:
    path = _write_cse(tmp_path / "cse.xlsx", "=IF(A1:A3>0,B1:B3,0)", ref="E1:E1")
    observed = _array_cell(path)
    graph = create_dependency_graph(path, ["Sheet1!E1"], load_values=False)
    node = graph.get_node("Sheet1!E1")
    assert node is not None
    assert node.is_array_formula is True
    assert node.array_formula_ref == observed.ref
    assert node.formula_ast is not None


def test_extract_stores_dynamic_array_spill_provenance(tmp_path: Path) -> None:
    path = _write_dynamic(tmp_path / "dyn.xlsx", "=A1:A3*2", ref="E1:E3")
    observed = _array_cell(path)
    graph = create_dependency_graph(path, ["Sheet1!E1"], load_values=False)
    node = graph.get_node("Sheet1!E1")
    assert node is not None
    assert node.is_array_formula is True
    assert node.array_formula_ref == observed.ref == "E1:E3"
    assert node.formula_ast is not None


def test_extract_does_not_mark_scalar_formula_as_array(tmp_path: Path) -> None:
    path = tmp_path / "scalar.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    ws.write_formula("B1", "=A1*2", None, 0)
    wb.close()
    graph = create_dependency_graph(path, ["Sheet1!B1"], load_values=False)
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    assert node.is_array_formula is False
    assert node.array_formula_ref is None


def test_write_workbook_round_trips_cse_array_formula(tmp_path: Path) -> None:
    source = _write_cse(tmp_path / "cse.xlsx", "=IF(A1:A3>0,B1:B3,0)", ref="E1:E1")
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(source, ["Sheet1!E1"], load_values=False)
    before = graph.get_node("Sheet1!E1")
    assert before is not None and before.formula_ast is not None
    observed = _array_cell(source)

    write_workbook(graph, dest)
    written = _array_cell(dest)
    assert written.ref == observed.ref
    assert written.text == render_formula(
        before.formula_ast,
        anchor=before.address,
        style=FormulaStyle.A1_EXCEL,
    )

    restored = create_dependency_graph(dest, ["Sheet1!E1"], load_values=False)
    after = restored.get_node("Sheet1!E1")
    assert after is not None
    assert after.is_array_formula is True
    assert after.array_formula_ref == before.array_formula_ref
    assert after.formula_ast == before.formula_ast


def test_write_workbook_round_trips_dynamic_array_spill(tmp_path: Path) -> None:
    source = _write_dynamic(tmp_path / "dyn.xlsx", "=A1:A3*2", ref="E1:E3")
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(source, ["Sheet1!E1"], load_values=False)
    before = graph.get_node("Sheet1!E1")
    assert before is not None and before.formula_ast is not None
    observed = _array_cell(source)

    write_workbook(graph, dest)
    written = _array_cell(dest)
    assert written.ref == observed.ref == "E1:E3"
    assert written.text == render_formula(
        before.formula_ast,
        anchor=before.address,
        style=FormulaStyle.A1_EXCEL,
    )

    restored = create_dependency_graph(dest, ["Sheet1!E1"], load_values=False)
    after = restored.get_node("Sheet1!E1")
    assert after is not None
    assert after.is_array_formula is True
    assert after.array_formula_ref == "E1:E3"
    assert after.formula_ast == before.formula_ast


def test_write_workbook_refuses_array_formula_without_ref(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "E",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=A1:A3*2", anchor="Sheet1!E1"),
            is_array_formula=True,
        )
    )
    with pytest.raises(ValueError, match="Sheet1!E1"):
        write_workbook(graph, tmp_path / "missing-ref.xlsx")
    assert not (tmp_path / "missing-ref.xlsx").exists()


def test_write_workbook_does_not_wrap_scalar_formula(tmp_path: Path) -> None:
    source = tmp_path / "scalar.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = xlsxwriter.Workbook(source)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    ws.write_formula("B1", "=A1*2", None, 0)
    wb.close()
    graph = create_dependency_graph(source, ["Sheet1!B1"], load_values=False)
    write_workbook(graph, dest)
    wb_out = fastpyxl.load_workbook(dest)
    try:
        raw = wb_out["Sheet1"]["B1"].value
    finally:
        wb_out.close()
    assert isinstance(raw, str)
    assert raw == "=A1*2"


def test_copy_node_preserves_array_formula_provenance() -> None:
    node = make_cell_node(
        "Sheet1",
        "E",
        1,
        is_leaf=False,
        formula_ast=parse_preserving_axes("=A1:A3*2", anchor="Sheet1!E1"),
        is_array_formula=True,
        array_formula_ref="E1:E3",
    )
    cloned = copy_node(node)
    assert cloned is not node
    assert cloned.is_array_formula is True
    assert cloned.array_formula_ref == "E1:E3"


def test_json_cache_round_trips_array_formula_provenance(tmp_path: Path) -> None:
    path = _write_dynamic(tmp_path / "dyn.xlsx", "=A1:A3*2", ref="E1:E3")
    graph = create_dependency_graph(path, ["Sheet1!E1"], load_values=False)
    assert GRAPH_CACHE_SCHEMA_VERSION >= 9
    restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    original = graph.get_node("Sheet1!E1")
    loaded = restored.get_node("Sheet1!E1")
    assert original is not None and loaded is not None
    assert loaded.is_array_formula is True
    assert loaded.array_formula_ref == original.array_formula_ref == "E1:E3"
    assert loaded.formula_ast == original.formula_ast
