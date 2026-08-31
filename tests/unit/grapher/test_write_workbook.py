"""`write_workbook` new-file extract round-trip (#564)."""

from __future__ import annotations

import typing
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import fastpyxl
import pytest

from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    FormulaStyle,
    FunctionCallNode,
    RangeNode,
    UnaryOpNode,
    parse_preserving_axes,
    render_formula,
    resolve_cell_ref,
)
from excel_grapher.core.types import XlError
from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.grapher.graph import DependencyGraph, GraphReadView
from excel_grapher.grapher.node import make_cell_node


def _write_mixed_source(path: Path) -> None:
    """Workbook with mixed `$` axes, a cross-sheet ref, a quoted sheet, and a stray cell."""
    wb = fastpyxl.Workbook()
    other = wb.active
    assert other is not None
    other.title = "Other"
    quoted = wb.create_sheet("Other Sheet")
    sheet1 = wb.create_sheet("Sheet1")

    other["B5"] = 10
    quoted["B5"] = 20
    sheet1["A1"] = 2
    sheet1["B2"] = "=$A$1+$A1+A$1+A1"
    sheet1["C1"] = "=Other!B5"
    sheet1["D1"] = "='Other Sheet'!B5"
    sheet1["Z99"] = 99
    wb.save(path)
    wb.close()


def _round_trip_targets() -> list[str]:
    return ["Sheet1!B2", "Sheet1!C1", "Sheet1!D1"]


def _ref_axes(ast: object, anchor: object) -> list[tuple[str, type, type]]:
    out: list[tuple[str, type, type]] = []

    def walk(node: object) -> None:
        match node:
            case CellRefNode(ref):
                out.append((resolve_cell_ref(ref, anchor), type(ref.col), type(ref.row)))
            case RangeNode(start_ref, end_ref):
                out.append(
                    (
                        resolve_cell_ref(start_ref, anchor),
                        type(start_ref.col),
                        type(start_ref.row),
                    )
                )
                out.append(
                    (
                        resolve_cell_ref(end_ref, anchor),
                        type(end_ref.col),
                        type(end_ref.row),
                    )
                )
            case FunctionCallNode(_, args):
                for arg in args:
                    walk(arg)
            case BinaryOpNode(_, left, right):
                walk(left)
                walk(right)
            case UnaryOpNode(_, operand):
                walk(operand)
            case _:
                return

    walk(ast)
    return out


@dataclass(frozen=True)
class _WrittenCell:
    value: object
    data_type: str
    cached_value: object | None


def _cell_at(path: Path, sheet: str, coord: str) -> _WrittenCell:
    wb = fastpyxl.load_workbook(path)
    try:
        cell = wb[sheet][coord]
        return _WrittenCell(cell.value, cell.data_type, getattr(cell, "cached_value", None))
    finally:
        wb.close()


def _sheetnames(path: Path) -> list[str]:
    wb = fastpyxl.load_workbook(path)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def test_write_workbook_graph_parameter_is_graph_read_view() -> None:
    hints = typing.get_type_hints(write_workbook)
    assert hints["graph"] is GraphReadView


def test_write_workbook_refuses_existing_destination_unless_overwrite(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_mixed_source(source)
    graph = create_dependency_graph(source, ["Sheet1!A1"])
    dest.write_bytes(b"sentinel")

    with pytest.raises(FileExistsError, match="out.xlsx"):
        write_workbook(graph, dest)
    assert dest.read_bytes() == b"sentinel"

    write_workbook(graph, dest, overwrite=True)
    wb = fastpyxl.load_workbook(dest)
    try:
        assert "Sheet1" in wb.sheetnames
        assert wb["Sheet1"]["A1"].value == 2
    finally:
        wb.close()


def test_write_workbook_refuses_empty_view(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="empty"):
        write_workbook(DependencyGraph(), tmp_path / "empty.xlsx")
    assert not (tmp_path / "empty.xlsx").exists()


def test_write_workbook_refuses_unparseable_formula(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "A",
            1,
            is_leaf=False,
            normalized_formula="=SUM(IF(@Sheet1!A1:A3>0,1,0))",
        )
    )
    with pytest.raises(ValueError, match="Sheet1!A1"):
        write_workbook(graph, tmp_path / "unparseable.xlsx")


def test_write_workbook_refuses_relative_axes_without_host_address(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    node = make_cell_node(
        "Sheet1",
        "B",
        2,
        is_leaf=False,
        formula_ast=parse_preserving_axes("=A1", anchor="Sheet1!B2"),
    )
    graph.add_node(node)
    graph._nodes[node.key].address = None
    with pytest.raises(ValueError, match="anchor"):
        write_workbook(graph, tmp_path / "no_anchor.xlsx")


def test_write_workbook_refuses_r1c1_for_normal_cells(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True))
    with pytest.raises(ValueError, match="R1C1"):
        write_workbook(
            graph,
            tmp_path / "r1c1.xlsx",
            formula_style=FormulaStyle.R1C1,
        )


def test_write_workbook_refuses_multi_sheet_without_sheet_order(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("B", "A", 1, value=1, is_leaf=True))
    graph.add_node(make_cell_node("A", "A", 1, value=2, is_leaf=True))
    with pytest.raises(ValueError, match="sheet_order"):
        write_workbook(graph, tmp_path / "no_order.xlsx")


def test_extract_write_extract_preserves_axis_kinds_and_omits_outside_cells(
    tmp_path: Path,
) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_mixed_source(source)
    graph = create_dependency_graph(
        source,
        _round_trip_targets(),
        store_raw_formula=True,
        load_values=True,
    )
    assert "Sheet1!Z99" not in graph

    write_workbook(graph, dest)
    written = create_dependency_graph(dest, _round_trip_targets(), load_values=True)

    assert set(written) == set(graph)
    assert _sheetnames(dest) == ["Other", "Other Sheet", "Sheet1"]

    for key in graph:
        before = graph.get_node(key)
        after = written.get_node(key)
        assert before is not None and after is not None
        if before.has_formula:
            assert before.formula_ast is not None
            assert after.formula_ast is not None
            assert _ref_axes(after.formula_ast, after.address) == _ref_axes(
                before.formula_ast, before.address
            )
            expected = render_formula(
                before.formula_ast,
                anchor=before.address,
                style=FormulaStyle.A1_EXCEL,
            )
            actual = render_formula(
                after.formula_ast,
                anchor=after.address,
                style=FormulaStyle.A1_EXCEL,
            )
            assert actual == expected
        else:
            assert after.value == before.value

    wb = fastpyxl.load_workbook(dest)
    try:
        assert wb["Sheet1"]["Z99"].value is None
        assert wb["Sheet1"]["B2"].value == "=$A$1+$A1+A$1+A1"
        assert wb["Sheet1"]["C1"].value == "=Other!B5"
        assert wb["Sheet1"]["D1"].value == "='Other Sheet'!B5"
    finally:
        wb.close()


def test_write_workbook_emits_rewritten_ast(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2
    ws["B1"] = "=A1+1"
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(source, ["Sheet1!B1"])
    graph.set_node_ast("Sheet1!B1", parse_preserving_axes("=A1*2", anchor="Sheet1!B1"))
    write_workbook(graph, dest)

    assert _cell_at(dest, "Sheet1", "B1").value == "=A1*2"
    rewritten = create_dependency_graph(dest, ["Sheet1!B1"])
    node = rewritten.get_node("Sheet1!B1")
    assert node is not None and node.formula_ast is not None
    assert (
        render_formula(node.formula_ast, anchor=node.address, style=FormulaStyle.A1_EXCEL)
        == "=A1*2"
    )


def test_write_workbook_does_not_write_formula_cached_values(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2
    ws["B1"] = "=A1+1"
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(source, ["Sheet1!B1"], load_values=True)
    graph.set_node_value("Sheet1!B1", 999)
    write_workbook(graph, dest)
    cell = _cell_at(dest, "Sheet1", "B1")
    assert cell.value == "=A1+1"
    assert getattr(cell, "cached_value", None) is None


def test_write_workbook_coerce_relative_refs_writes_absolute_a1(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_mixed_source(source)
    graph = create_dependency_graph(source, ["Sheet1!B2"])
    write_workbook(graph, dest, coerce_relative_refs=True)
    assert _cell_at(dest, "Sheet1", "B2").value == "=$A$1+$A$1+$A$1+$A$1"


def test_write_workbook_projection_omits_collapsed_addresses(tmp_path: Path) -> None:
    from excel_grapher.exporter import IdentityTransitCompression

    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    engine = wb.active
    assert engine is not None
    engine.title = "Engine"
    outputs = wb.create_sheet("Outputs")
    engine["C6"] = 10
    outputs["B12"] = "=Engine!C6"
    outputs["B14"] = "=Outputs!B12+1"
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(
        source,
        ["Outputs!B14"],
        capture_dependency_provenance=True,
        load_values=True,
    )
    assert "Outputs!B12" in graph
    original_keys = set(graph)
    projection = IdentityTransitCompression().project(graph)
    assert isinstance(projection, GraphReadView)
    assert "Outputs!B12" not in projection
    assert "Outputs!B12" in graph

    write_workbook(projection, dest)
    assert set(graph) == original_keys

    written = create_dependency_graph(dest, ["Outputs!B14"])
    assert "Outputs!B12" not in written
    assert "Outputs!B14" in written
    assert "Engine!C6" in written
    node = written.get_node("Outputs!B14")
    assert node is not None and node.formula_ast is not None
    assert (
        render_formula(node.formula_ast, anchor=node.address, style=FormulaStyle.A1_EXCEL)
        == "=Engine!C6+1"
    )
    wb_out = fastpyxl.load_workbook(dest)
    try:
        assert wb_out["Outputs"]["B12"].value is None
    finally:
        wb_out.close()


def test_write_workbook_leaf_types_round_trip(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2
    ws["A2"] = "hello"
    ws["A3"] = True
    ws["A4"] = None
    ws["A5"] = datetime(2024, 6, 15, 12, 30)
    ws["A6"] = "#DIV/0!"
    ws["B1"] = "=A1+A3"
    wb.save(source)
    wb.close()

    graph = create_dependency_graph(
        source,
        ["Sheet1!B1", "Sheet1!A2", "Sheet1!A4", "Sheet1!A5", "Sheet1!A6"],
        load_values=True,
    )
    write_workbook(graph, dest)
    written = create_dependency_graph(
        dest,
        ["Sheet1!B1", "Sheet1!A2", "Sheet1!A4", "Sheet1!A5", "Sheet1!A6"],
        load_values=True,
    )

    def value(key: str) -> object:
        node = written.get_node(key)
        assert node is not None
        return node.value

    assert value("Sheet1!A1") == 2
    assert value("Sheet1!A2") == "hello"
    assert value("Sheet1!A3") is True
    assert value("Sheet1!A4") is None
    assert value("Sheet1!A5") == datetime(2024, 6, 15, 12, 30)
    assert value("Sheet1!A6") in ("#DIV/0!", XlError.DIV)
    wb_out = fastpyxl.load_workbook(dest)
    try:
        assert wb_out["Sheet1"]["A6"].data_type == "e"
    finally:
        wb_out.close()


def test_write_workbook_accepts_xlerror_leaf(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=XlError.DIV, is_leaf=True))
    dest = tmp_path / "err.xlsx"
    write_workbook(graph, dest)
    cell = _cell_at(dest, "Sheet1", "A1")
    assert cell.value == "#DIV/0!"
    assert cell.data_type == "e"


def test_write_workbook_appends_sheets_missing_from_sheet_order(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Known"])
    graph.add_node(make_cell_node("Known", "A", 1, value=1, is_leaf=True))
    graph.add_node(make_cell_node("Zeta", "A", 1, value=2, is_leaf=True))
    graph.add_node(make_cell_node("Alpha", "A", 1, value=3, is_leaf=True))
    dest = tmp_path / "order.xlsx"
    write_workbook(graph, dest)
    wb = fastpyxl.load_workbook(dest)
    try:
        assert wb.sheetnames == ["Known", "Alpha", "Zeta"]
    finally:
        wb.close()
