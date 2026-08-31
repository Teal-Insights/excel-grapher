"""Shared-formula R1C1 grouping on `write_workbook` (#567)."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
import xlsxwriter

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
from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.grapher.formula_shapes import warm_formula_shapes
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node

_SSML = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _write_autofill_column(path: Path, *, rows: int = 5, gap_at: int | None = None) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for i in range(rows):
        ws.write_number(i, 0, i + 1)  # A1:A{rows}
        if gap_at is not None and i + 1 == gap_at:
            ws.write_number(i, 1, 99)
        else:
            ws.write_formula(i, 1, f"=A{i + 1}+1", None, i + 2)
    wb.close()
    return path


def _write_autofill_row(path: Path, *, cols: int = 4) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for i in range(cols):
        ws.write_number(0, i, i + 1)  # A1:D1
        ws.write_formula(1, i, f"={chr(ord('A') + i)}1*2", None, (i + 1) * 2)
    wb.close()
    return path


def _shared_formula_cells(path: Path) -> list[tuple[str, dict[str, str], str | None]]:
    with zipfile.ZipFile(path) as zf:
        name = next(n for n in zf.namelist() if n.startswith("xl/worksheets/sheet"))
        xml = zf.read(name)
    root = ET.fromstring(xml)
    found: list[tuple[str, dict[str, str], str | None]] = []
    for cell in root.findall(".//m:c", _SSML):
        formula = cell.find("m:f", _SSML)
        if formula is None or formula.get("t") != "shared":
            continue
        found.append((str(cell.get("r")), dict(formula.attrib), formula.text))
    return found


def _scalar_formula_cells(path: Path) -> list[tuple[str, str | None]]:
    with zipfile.ZipFile(path) as zf:
        name = next(n for n in zf.namelist() if n.startswith("xl/worksheets/sheet"))
        xml = zf.read(name)
    root = ET.fromstring(xml)
    found: list[tuple[str, str | None]] = []
    for cell in root.findall(".//m:c", _SSML):
        formula = cell.find("m:f", _SSML)
        if formula is None:
            continue
        if formula.get("t") in {"shared", "array", "dataTable"}:
            continue
        found.append((str(cell.get("r")), formula.text))
    return found


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


def test_write_workbook_groups_autofill_column_as_shared_formula(tmp_path: Path) -> None:
    source = _write_autofill_column(tmp_path / "src.xlsx")
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(
        source,
        ["Sheet1!B1:B5"],
        load_values=False,
        warm_formula_shapes=True,
    )
    keys = ("Sheet1!B1", "Sheet1!B2", "Sheet1!B3", "Sheet1!B4", "Sheet1!B5")
    r1c1 = {}
    for key in keys:
        node = graph.get_node(key)
        assert node is not None and node.formula_ast is not None
        r1c1[key] = render_formula(node.formula_ast, anchor=node.address, style=FormulaStyle.R1C1)
    assert len(set(r1c1.values())) == 1
    assert next(iter(r1c1.values())) == "=RC[-1]+1"

    write_workbook(graph, dest)

    shared = _shared_formula_cells(dest)
    coords = {coord for coord, _attrs, _text in shared}
    assert coords == {"B1", "B2", "B3", "B4", "B5"}
    masters = [(coord, attrs, text) for coord, attrs, text in shared if "ref" in attrs]
    assert len(masters) == 1
    master_coord, master_attrs, master_text = masters[0]
    assert master_coord == "B1"
    assert master_attrs["t"] == "shared"
    assert master_attrs["si"] == "0"
    assert master_attrs["ref"] == "B1:B5"
    assert master_text == "A1+1"
    for coord, attrs, text in shared:
        if coord == master_coord:
            continue
        assert attrs == {"t": "shared", "si": "0"}
        assert text is None

    restored = create_dependency_graph(dest, ["Sheet1!B1:B5"], load_values=False)
    for key in r1c1:
        before = graph.get_node(key)
        after = restored.get_node(key)
        assert before is not None and after is not None
        assert before.formula_ast is not None and after.formula_ast is not None
        assert _ref_axes(after.formula_ast, after.address) == _ref_axes(
            before.formula_ast, before.address
        )
        assert render_formula(
            after.formula_ast, anchor=after.address, style=FormulaStyle.A1_EXCEL
        ) == render_formula(before.formula_ast, anchor=before.address, style=FormulaStyle.A1_EXCEL)
        assert (
            render_formula(after.formula_ast, anchor=after.address, style=FormulaStyle.R1C1)
            == "=RC[-1]+1"
        )


def test_write_workbook_groups_autofill_row_as_shared_formula(tmp_path: Path) -> None:
    source = _write_autofill_row(tmp_path / "row.xlsx")
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(
        source,
        ["Sheet1!A2:D2"],
        load_values=False,
        warm_formula_shapes=True,
    )
    write_workbook(graph, dest)
    shared = _shared_formula_cells(dest)
    coords = {coord for coord, _attrs, _text in shared}
    assert coords == {"A2", "B2", "C2", "D2"}
    masters = [item for item in shared if "ref" in item[1]]
    assert len(masters) == 1
    assert masters[0][1]["ref"] == "A2:D2"
    assert masters[0][2] == "A1*2"


def test_write_workbook_skips_shared_formulas_when_shapes_are_cold(tmp_path: Path) -> None:
    source = _write_autofill_column(tmp_path / "src.xlsx")
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(source, ["Sheet1!B1:B5"], load_values=False)
    assert graph.formula_shapes is None
    write_workbook(graph, dest)
    assert _shared_formula_cells(dest) == []
    scalars = dict(_scalar_formula_cells(dest))
    assert scalars["B1"] == "A1+1"
    assert scalars["B5"] == "A5+1"


def test_write_workbook_require_shared_formulas_fails_when_shapes_missing(
    tmp_path: Path,
) -> None:
    source = _write_autofill_column(tmp_path / "src.xlsx", rows=2)
    graph = create_dependency_graph(source, ["Sheet1!B1:B2"], load_values=False)
    with pytest.raises(ValueError, match="formula_shapes"):
        write_workbook(graph, tmp_path / "out.xlsx", shared_formulas="require")
    assert not (tmp_path / "out.xlsx").exists()


def test_write_workbook_shared_formulas_off_emits_per_cell(tmp_path: Path) -> None:
    source = _write_autofill_column(tmp_path / "src.xlsx")
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(
        source,
        ["Sheet1!B1:B5"],
        load_values=False,
        warm_formula_shapes=True,
    )
    write_workbook(graph, dest, shared_formulas="off")
    assert _shared_formula_cells(dest) == []
    assert dict(_scalar_formula_cells(dest))["B3"] == "A3+1"


def test_write_workbook_does_not_group_noncontiguous_autofill(tmp_path: Path) -> None:
    source = _write_autofill_column(tmp_path / "gap.xlsx", rows=5, gap_at=3)
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(
        source,
        ["Sheet1!B1:B5"],
        load_values=False,
        warm_formula_shapes=True,
    )
    write_workbook(graph, dest)
    shared = _shared_formula_cells(dest)
    refs = {attrs["ref"] for _coord, attrs, _text in shared if "ref" in attrs}
    assert refs == {"B1:B2", "B4:B5"}
    coords = {coord for coord, _attrs, _text in shared}
    assert coords == {"B1", "B2", "B4", "B5"}
    assert "B3" not in coords


def test_write_workbook_does_not_group_distinct_absolute_params(tmp_path: Path) -> None:
    source = tmp_path / "abs.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = xlsxwriter.Workbook(source)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    ws.write_number(1, 0, 2)
    ws.write_formula(0, 1, "=$A$1+1", None, 2)
    ws.write_formula(1, 1, "=$A$2+1", None, 3)
    wb.close()
    graph = create_dependency_graph(
        source,
        ["Sheet1!B1:B2"],
        load_values=False,
        warm_formula_shapes=True,
    )
    write_workbook(graph, dest)
    assert _shared_formula_cells(dest) == []
    scalars = dict(_scalar_formula_cells(dest))
    assert scalars["B1"] == "$A$1+1"
    assert scalars["B2"] == "$A$2+1"


def test_write_workbook_does_not_group_stale_formula_shapes(tmp_path: Path) -> None:
    source = _write_autofill_column(tmp_path / "src.xlsx", rows=3)
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(
        source,
        ["Sheet1!B1:B3"],
        load_values=False,
        warm_formula_shapes=True,
    )
    stale = graph.formula_shapes
    assert stale is not None
    graph.set_node_ast("Sheet1!B2", parse_preserving_axes("=A2*10", anchor="Sheet1!B2"))
    graph.formula_shapes = stale
    write_workbook(graph, dest)
    shared = _shared_formula_cells(dest)
    coords = {coord for coord, _attrs, _text in shared}
    assert "B2" not in coords
    scalars = dict(_scalar_formula_cells(dest))
    assert scalars["B2"] == "A2*10"


def test_write_workbook_does_not_group_when_coerce_relative_refs(tmp_path: Path) -> None:
    source = _write_autofill_column(tmp_path / "src.xlsx", rows=3)
    dest = tmp_path / "out.xlsx"
    graph = create_dependency_graph(
        source,
        ["Sheet1!B1:B3"],
        load_values=False,
        warm_formula_shapes=True,
    )
    write_workbook(graph, dest, coerce_relative_refs=True)
    assert _shared_formula_cells(dest) == []
    scalars = dict(_scalar_formula_cells(dest))
    assert scalars["B1"] == "$A$1+1"
    assert scalars["B2"] == "$A$2+1"


def test_write_workbook_does_not_group_indirect(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True))
    graph.add_node(make_cell_node("Sheet1", "A", 2, value=2, is_leaf=True))
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            1,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=INDIRECT(A1)", anchor="Sheet1!B1"),
        )
    )
    graph.add_node(
        make_cell_node(
            "Sheet1",
            "B",
            2,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=INDIRECT(A2)", anchor="Sheet1!B2"),
        )
    )
    graph.formula_shapes = warm_formula_shapes(graph)
    dest = tmp_path / "indirect.xlsx"
    write_workbook(graph, dest)
    assert _shared_formula_cells(dest) == []


def test_write_workbook_groups_mixed_relative_and_absolute_autofill(tmp_path: Path) -> None:
    source = tmp_path / "mixed.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = xlsxwriter.Workbook(source)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 10)
    for i in range(3):
        ws.write_number(i, 1, i + 1)
        ws.write_formula(i, 2, f"=B{i + 1}+$A$1", None, i + 11)
    wb.close()
    graph = create_dependency_graph(
        source,
        ["Sheet1!C1:C3"],
        load_values=False,
        warm_formula_shapes=True,
    )
    write_workbook(graph, dest)
    shared = _shared_formula_cells(dest)
    refs = {attrs["ref"] for _coord, attrs, _text in shared if "ref" in attrs}
    assert refs == {"C1:C3"}
    masters = [text for _coord, attrs, text in shared if "ref" in attrs]
    assert masters == ["B1+$A$1"]
    restored = create_dependency_graph(dest, ["Sheet1!C1:C3"], load_values=False)
    for key in ("Sheet1!C1", "Sheet1!C2", "Sheet1!C3"):
        before = graph.get_node(key)
        after = restored.get_node(key)
        assert before is not None and after is not None
        assert before.formula_ast is not None and after.formula_ast is not None
        assert _ref_axes(after.formula_ast, after.address) == _ref_axes(
            before.formula_ast, before.address
        )


def test_write_workbook_rejects_unknown_shared_formulas_mode(tmp_path: Path) -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    graph.add_node(make_cell_node("Sheet1", "A", 1, value=1, is_leaf=True))
    with pytest.raises(ValueError, match="shared_formulas"):
        write_workbook(graph, tmp_path / "out.xlsx", shared_formulas="sometimes")  # type: ignore[arg-type]
    assert not (tmp_path / "out.xlsx").exists()
