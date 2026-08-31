"""`write_workbook` round-trip after `move_node` geometry edits (#566)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    BinaryOpNode,
    CellRefNode,
    FormulaStyle,
    FunctionCallNode,
    RangeNode,
    RelativeAxis,
    UnaryOpNode,
    render_formula,
    resolve_cell_ref,
)
from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.grapher.node import NodeView

_MIXED_AXIS_KINDS = (
    (AbsoluteAxis, AbsoluteAxis),
    (AbsoluteAxis, RelativeAxis),
    (RelativeAxis, AbsoluteAxis),
    (RelativeAxis, RelativeAxis),
)


def _write_host_formula(path: Path, formula: str) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2
    ws["B2"] = formula
    wb.save(path)
    wb.close()


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


def _xlsx_value(path: Path, sheet: str, coord: str) -> object:
    wb = fastpyxl.load_workbook(path)
    try:
        return wb[sheet][coord].value
    finally:
        wb.close()


def _a1_excel(node: NodeView) -> str:
    assert node.formula_ast is not None
    return render_formula(
        node.formula_ast,
        anchor=node.address,
        style=FormulaStyle.A1_EXCEL,
    )


def test_write_workbook_after_host_move_keeps_resolved_target(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_host_formula(source, "=A1")
    graph = create_dependency_graph(source, ["Sheet1!B2"], load_values=True)

    graph.move_node("Sheet1!B2", "Sheet1!C3")
    write_workbook(graph, dest)

    assert _xlsx_value(dest, "Sheet1", "C3") == "=A1"
    assert _xlsx_value(dest, "Sheet1", "B2") is None

    written = create_dependency_graph(dest, ["Sheet1!C3"], load_values=True)
    assert "Sheet1!B2" not in written
    moved = written.get_node("Sheet1!C3")
    assert moved is not None and moved.formula_ast is not None
    assert isinstance(moved.formula_ast, CellRefNode)
    assert moved.formula_ast.ref.col == RelativeAxis(-2)
    assert moved.formula_ast.ref.row == RelativeAxis(-2)
    assert resolve_cell_ref(moved.formula_ast.ref, moved.address) == "Sheet1!A1"
    assert _a1_excel(moved) == "=A1"
    leaf = written.get_node("Sheet1!A1")
    assert leaf is not None and leaf.value == 2


def test_write_workbook_after_referenced_cell_move_retargets_dependents(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_host_formula(source, "=A1")
    graph = create_dependency_graph(source, ["Sheet1!B2"], load_values=True)

    graph.move_node("Sheet1!A1", "Sheet1!C3")
    write_workbook(graph, dest)

    assert _xlsx_value(dest, "Sheet1", "B2") == "=C3"
    assert _xlsx_value(dest, "Sheet1", "C3") == 2
    assert _xlsx_value(dest, "Sheet1", "A1") is None

    written = create_dependency_graph(dest, ["Sheet1!B2"], load_values=True)
    assert "Sheet1!A1" not in written
    host = written.get_node("Sheet1!B2")
    assert host is not None and host.formula_ast is not None
    assert isinstance(host.formula_ast, CellRefNode)
    assert host.formula_ast.ref.col == RelativeAxis(1)
    assert host.formula_ast.ref.row == RelativeAxis(1)
    assert resolve_cell_ref(host.formula_ast.ref, host.address) == "Sheet1!C3"
    assert _a1_excel(host) == "=C3"
    leaf = written.get_node("Sheet1!C3")
    assert leaf is not None and leaf.value == 2


def test_write_workbook_after_host_move_preserves_mixed_dollar_axes(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    formula = "=$A$1+$A1+A$1+A1"
    _write_host_formula(source, formula)
    graph = create_dependency_graph(source, ["Sheet1!B2"], load_values=True)
    graph.move_node("Sheet1!B2", "Sheet1!C3")
    write_workbook(graph, dest)

    assert _xlsx_value(dest, "Sheet1", "C3") == formula
    assert _xlsx_value(dest, "Sheet1", "B2") is None
    written = create_dependency_graph(dest, ["Sheet1!C3"])
    moved = written.get_node("Sheet1!C3")
    assert moved is not None and moved.formula_ast is not None
    axes = _ref_axes(moved.formula_ast, moved.address)
    assert [target for target, _, _ in axes] == ["Sheet1!A1"] * 4
    assert [(col, row) for _, col, row in axes] == list(_MIXED_AXIS_KINDS)
    assert _a1_excel(moved) == formula


def test_write_workbook_after_referenced_move_preserves_mixed_dollar_axes(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_host_formula(source, "=$A$1+$A1+A$1+A1")
    graph = create_dependency_graph(source, ["Sheet1!B2"], load_values=True)
    graph.move_node("Sheet1!A1", "Sheet1!D5")
    write_workbook(graph, dest)

    assert _xlsx_value(dest, "Sheet1", "B2") == "=$D$5+$D5+D$5+D5"
    assert _xlsx_value(dest, "Sheet1", "D5") == 2
    assert _xlsx_value(dest, "Sheet1", "A1") is None
    written = create_dependency_graph(dest, ["Sheet1!B2"])
    dependent = written.get_node("Sheet1!B2")
    assert dependent is not None and dependent.formula_ast is not None
    axes = _ref_axes(dependent.formula_ast, dependent.address)
    assert [target for target, _, _ in axes] == ["Sheet1!D5"] * 4
    assert [(col, row) for _, col, row in axes] == list(_MIXED_AXIS_KINDS)
    assert _a1_excel(dependent) == "=$D$5+$D5+D$5+D5"


def test_write_workbook_stale_projection_does_not_see_later_move(tmp_path: Path) -> None:
    from excel_grapher.exporter import IdentityTransitCompression

    source = tmp_path / "src.xlsx"
    dest_stale = tmp_path / "stale-projection.xlsx"
    dest_moved = tmp_path / "moved.xlsx"
    dest_reprojected = tmp_path / "reprojected.xlsx"

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
    projection = IdentityTransitCompression().project(graph)
    assert "Outputs!B12" not in projection
    assert "Outputs!B14" in projection

    graph.move_node("Outputs!B14", "Outputs!C20")
    assert "Outputs!C20" in graph
    assert "Outputs!B14" not in graph
    assert "Outputs!B14" in projection
    assert "Outputs!C20" not in projection

    write_workbook(projection, dest_stale)
    stale = create_dependency_graph(dest_stale, ["Outputs!B14"])
    assert "Outputs!B14" in stale
    assert "Outputs!C20" not in stale
    assert "Outputs!B12" not in stale
    assert _xlsx_value(dest_stale, "Outputs", "B14") == "=Engine!C6+1"
    assert _xlsx_value(dest_stale, "Outputs", "C20") is None

    write_workbook(graph, dest_moved)
    moved = create_dependency_graph(dest_moved, ["Outputs!C20"])
    assert "Outputs!C20" in moved
    assert "Outputs!B14" not in moved
    assert "Outputs!B12" in moved
    assert _xlsx_value(dest_moved, "Outputs", "C20") == "=B12+1"
    assert _xlsx_value(dest_moved, "Outputs", "B14") is None

    reprojected = IdentityTransitCompression().project(graph)
    write_workbook(reprojected, dest_reprojected)
    after = create_dependency_graph(dest_reprojected, ["Outputs!C20"])
    assert "Outputs!C20" in after
    assert "Outputs!B14" not in after
    assert "Outputs!B12" not in after
    node = after.get_node("Outputs!C20")
    assert node is not None
    assert _a1_excel(node) == "=Engine!C6+1"
    assert _xlsx_value(dest_reprojected, "Outputs", "B12") is None
    assert _xlsx_value(dest_reprojected, "Outputs", "B14") is None
