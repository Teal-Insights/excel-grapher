"""Axis-level relative/absolute cell refs (#545)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.address_keys import CellKey
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    BinaryOpNode,
    CellRef,
    CellRefNode,
    FormulaParseError,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    RelativeAxis,
    WholeColumnNode,
    WholeRowNode,
    iter_resolved_cell_keys,
    parse,
    parse_formula_text,
    parse_preserving_axes,
    parse_preserving_axes_optional,
    resolve_cell_ref,
)
from excel_grapher.core.formula_shape import (
    fingerprint_formula_shape,
    intern_formula_shapes,
    resolve_address_leaf,
)
from excel_grapher.grapher.formula_shapes import warm_formula_shapes
from excel_grapher.grapher.parser import parse_dynamic_range_refs_with_spans


def _rel(col: int, row: int, *, sheet: str = "Sheet1") -> CellRef:
    return CellRef(sheet=sheet, col=RelativeAxis(col), row=RelativeAxis(row))


def _abs(col: int, row: int, *, sheet: str = "Sheet1") -> CellRef:
    return CellRef(sheet=sheet, col=AbsoluteAxis(col), row=AbsoluteAxis(row))


def test_resolve_cell_ref_absolute_ignores_anchor() -> None:
    ref = _abs(1, 1)
    assert resolve_cell_ref(ref, "Sheet1!Z99") == "Sheet1!A1"
    assert resolve_cell_ref(ref, None) == "Sheet1!A1"


def test_resolve_cell_ref_relative_requires_anchor() -> None:
    ref = _rel(-1, 0)
    with pytest.raises(ValueError, match="anchor"):
        resolve_cell_ref(ref, None)
    assert resolve_cell_ref(ref, "Sheet1!B2") == "Sheet1!A2"
    assert resolve_cell_ref(ref, CellKey("Sheet1!C5")) == "Sheet1!B5"


def test_resolve_cell_ref_mixed_axes() -> None:
    ref = CellRef(sheet="Sheet1", col=AbsoluteAxis(1), row=RelativeAxis(-1))
    assert resolve_cell_ref(ref, "Sheet1!B3") == "Sheet1!A2"


def test_parse_normalized_formula_is_fully_absolute() -> None:
    node = parse("=Sheet1!A1")
    assert isinstance(node, CellRefNode)
    assert node.ref == _abs(1, 1)
    assert node.address == "Sheet1!A1"
    assert parse("=Sheet1!$A$1") == node


def test_parse_preserving_axes_relative_offset_from_anchor() -> None:
    ast = parse_preserving_axes("=A1+1", anchor="Sheet1!B2")
    assert ast == BinaryOpNode("+", CellRefNode(_rel(-1, -1)), NumberNode(1.0))


def test_parse_preserving_axes_absolute_dollar_markers() -> None:
    ast = parse_preserving_axes("=$A$1", anchor="Sheet1!B2")
    assert ast == CellRefNode(_abs(1, 1))


def test_parse_preserving_axes_mixed_row_relative() -> None:
    ast = parse_preserving_axes("=$A1", anchor="Sheet1!B2")
    assert ast == CellRefNode(CellRef(sheet="Sheet1", col=AbsoluteAxis(1), row=RelativeAxis(-1)))


def test_parse_preserving_axes_mixed_col_relative() -> None:
    ast = parse_preserving_axes("=A$1", anchor="Sheet1!B2")
    assert ast == CellRefNode(CellRef(sheet="Sheet1", col=RelativeAxis(-1), row=AbsoluteAxis(1)))


def test_parse_preserving_axes_cross_sheet_keeps_sheet_and_axes() -> None:
    ast = parse_preserving_axes("=Data!B3", anchor="Sheet1!C5")
    assert ast == CellRefNode(_rel(-1, -2, sheet="Data"))
    ast_abs = parse_preserving_axes("=Data!$B$3", anchor="Sheet1!C5")
    assert ast_abs == CellRefNode(_abs(2, 3, sheet="Data"))


def test_parse_preserving_axes_quoted_sheet() -> None:
    ast = parse_preserving_axes("='My Sheet'!A1", anchor="'My Sheet'!B2")
    assert isinstance(ast, CellRefNode)
    assert ast == CellRefNode(_rel(-1, -1, sheet="My Sheet"))
    assert resolve_cell_ref(ast.ref, "'My Sheet'!B2") == "'My Sheet'!A1"


def test_parse_preserving_axes_range_endpoints() -> None:
    ast = parse_preserving_axes("=SUM(A1:A3)", anchor="Sheet1!B4")
    assert isinstance(ast, FunctionCallNode)
    rng = ast.args[0]
    assert isinstance(rng, RangeNode)
    assert rng.start_ref == _rel(-1, -3)
    assert rng.end_ref == _rel(-1, -1)
    keys = list(iter_resolved_cell_keys(ast, "Sheet1!B4"))
    assert keys == ["Sheet1!A1", "Sheet1!A3"]


def test_parse_preserving_axes_whole_column_relative_vs_absolute() -> None:
    rel = parse_preserving_axes("=A:A", anchor="Sheet1!B1")
    assert rel == WholeColumnNode(sheet="Sheet1", col=RelativeAxis(-1))
    abs_col = parse_preserving_axes("=$A:$A", anchor="Sheet1!B1")
    assert isinstance(abs_col, WholeColumnNode)
    assert abs_col == WholeColumnNode(sheet="Sheet1", col=AbsoluteAxis(1))
    assert abs_col.column == "A"


def test_parse_preserving_axes_whole_row_relative_vs_absolute() -> None:
    rel = parse_preserving_axes("=1:1", anchor="Sheet1!A3")
    assert rel == WholeRowNode(sheet="Sheet1", row=RelativeAxis(-2))
    abs_row = parse_preserving_axes("=$1:$1", anchor="Sheet1!A3")
    assert abs_row == WholeRowNode(sheet="Sheet1", row=AbsoluteAxis(1))


def test_parse_preserving_axes_named_range_is_absolute() -> None:
    ast = parse_preserving_axes(
        "=MyName+A1",
        anchor="Sheet1!B2",
        named_ranges={"MyName": ("Sheet1", "Z9")},
    )
    assert ast == BinaryOpNode(
        "+",
        CellRefNode(_abs(26, 9)),
        CellRefNode(_rel(-1, -1)),
    )


def test_parse_preserving_axes_optional_fail_soft() -> None:
    assert parse_preserving_axes_optional("=SUM(IF(@A1>0,1,0))", anchor="Sheet1!B1") is None
    assert parse_preserving_axes_optional(None, anchor="Sheet1!B1") is None


def test_parse_formula_text_preserves_axes_when_anchored() -> None:
    ast = parse_formula_text("=A1+$B$1", anchor="Sheet1!B2")
    assert ast == parse_preserving_axes("=A1+$B$1", anchor="Sheet1!B2")
    assert parse_formula_text("=SUM(IF(@A1>0,1,0))", anchor="Sheet1!B1") is None
    assert parse_formula_text(None, anchor="Sheet1!B1") is None


def test_parse_formula_text_absolutizes_without_anchor() -> None:
    ast = parse_formula_text("=Sheet1!A1")
    assert ast == parse("=Sheet1!A1")
    assert isinstance(ast, CellRefNode)
    assert isinstance(ast.ref.col, AbsoluteAxis)
    assert isinstance(ast.ref.row, AbsoluteAxis)


def test_parse_preserving_axes_accepts_excel_like_whitespace() -> None:
    ast = parse_preserving_axes("=SUM( A1 )", anchor="Sheet1!B1")
    assert ast == FunctionCallNode("SUM", [CellRefNode(_rel(-1, 0))])
    spaced = parse_preserving_axes("= A1 + 1.0 ", anchor="Sheet1!B2")
    assert spaced == BinaryOpNode("+", CellRefNode(_rel(-1, -1)), NumberNode(1.0))


def test_parse_preserving_axes_scientific_literals() -> None:
    ast = parse_preserving_axes("=A1+1e2", anchor="Sheet1!B1")
    assert ast == BinaryOpNode("+", CellRefNode(_rel(-1, 0)), NumberNode(100.0))
    assert parse_preserving_axes_optional("=1e", anchor="Sheet1!A1") is None


def test_parse_preserving_axes_rejects_r1c1_brackets() -> None:
    with pytest.raises(FormulaParseError):
        parse_preserving_axes("=R[-1]C[-1]", anchor="Sheet1!B2")


def test_indirect_r1c1_stays_fail_closed() -> None:
    with pytest.raises(ValueError, match="R1C1"):
        parse_dynamic_range_refs_with_spans(
            '=INDIRECT("R[-1]C[-1]", FALSE)',
            current_sheet="Sheet1",
            current_cell_a1="B2",
        )


def test_resolve_address_leaf_binds_relative_params_to_host() -> None:
    rel = CellRefNode(_rel(-1, 0))
    assert resolve_address_leaf(rel, "Sheet1!B2") == "Sheet1!A2"
    assert resolve_address_leaf(CellRefNode(_abs(1, 1)), "Sheet1!Z99") == "Sheet1!A1"
    rng = RangeNode(_rel(-1, -1), _rel(-1, 0))
    assert resolve_address_leaf(rng, "Sheet1!B3") == "Sheet1!A2:A3"
    col = WholeColumnNode(sheet="Sheet1", col=RelativeAxis(-1))
    assert resolve_address_leaf(col, "Sheet1!B1") == "Sheet1!A:A"
    row = WholeRowNode(sheet="Sheet1", row=RelativeAxis(-1))
    assert resolve_address_leaf(row, "Sheet1!A3") == "Sheet1!2:2"


def test_autofill_siblings_share_relative_shape_params() -> None:
    left = parse_preserving_axes("=A1*2", anchor="Sheet1!B1")
    right = parse_preserving_axes("=A2*2", anchor="Sheet1!B2")
    a = fingerprint_formula_shape(left)
    b = fingerprint_formula_shape(right)
    assert a.shape_key == b.shape_key
    assert a.params == b.params
    assert a.params == (CellRefNode(_rel(-1, 0)),)

    mixed_a = parse_preserving_axes("=$A1+1", anchor="Sheet1!B2")
    mixed_b = parse_preserving_axes("=$A2+1", anchor="Sheet1!C3")
    assert fingerprint_formula_shape(mixed_a).params == fingerprint_formula_shape(mixed_b).params


def test_intern_formula_shapes_collapses_relative_autofill_params() -> None:
    table = intern_formula_shapes(
        [
            ("Sheet1!B1", parse_preserving_axes("=A1*2", anchor="Sheet1!B1")),
            ("Sheet1!B2", parse_preserving_axes("=A2*2", anchor="Sheet1!B2")),
            ("Sheet1!C1", parse_preserving_axes("=$A$1*2", anchor="Sheet1!C1")),
        ]
    )
    b1 = table.lookup("Sheet1!B1")
    b2 = table.lookup("Sheet1!B2")
    c1 = table.lookup("Sheet1!C1")
    assert b1 is not None and b2 is not None and c1 is not None
    assert b1[0] == b2[0] == c1[0]
    assert b1[2] == b2[2]
    assert b1[2] != c1[2]


def _autofill_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "axis_autofill.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["A2"].value = 20
    ws["B1"].value = "=A1*2"
    ws["B2"].value = "=A2*2"
    ws["C1"].value = "=$A$1*2"
    wb.save(path)
    wb.close()
    return path


def test_extraction_stores_relative_ast_and_absolute_normalized_formula(tmp_path: Path) -> None:
    path = _autofill_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2", "Sheet1!C1"], load_values=True)
    b1 = graph._get_internal_node("Sheet1!B1")
    b2 = graph._get_internal_node("Sheet1!B2")
    c1 = graph._get_internal_node("Sheet1!C1")
    assert b1 is not None and b2 is not None and c1 is not None
    assert b1.normalized_formula == "=Sheet1!A1*2"
    assert b2.normalized_formula == "=Sheet1!A2*2"
    assert c1.normalized_formula == "=Sheet1!A1*2"
    assert b1.formula_ast == parse_preserving_axes("=A1*2", anchor="Sheet1!B1")
    assert b1.formula_ast == b2.formula_ast
    assert b1.formula_ast is b2.formula_ast
    assert c1.formula_ast != b1.formula_ast


def test_extraction_shape_params_match_for_autofill_siblings(tmp_path: Path) -> None:
    path = _autofill_workbook(tmp_path)
    graph = create_dependency_graph(
        path,
        ["Sheet1!B1", "Sheet1!B2", "Sheet1!C1"],
        load_values=False,
        warm_formula_shapes=True,
    )
    table = warm_formula_shapes(graph) if graph.formula_shapes is None else graph.formula_shapes
    b1 = table.lookup("Sheet1!B1")
    b2 = table.lookup("Sheet1!B2")
    c1 = table.lookup("Sheet1!C1")
    assert b1 is not None and b2 is not None and c1 is not None
    assert b1[2] == b2[2]
    assert b1[2] != c1[2]


def test_evaluator_resolves_relative_refs_against_host_cell(tmp_path: Path) -> None:
    path = _autofill_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B1", "Sheet1!B2", "Sheet1!C1"], load_values=True)
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["Sheet1!B1"])["Sheet1!B1"] == 20.0
        assert ev.evaluate(["Sheet1!B2"])["Sheet1!B2"] == 40.0
        assert ev.evaluate(["Sheet1!C1"])["Sheet1!C1"] == 20.0


def test_evaluator_with_shape_helpers_resolves_relative_params(tmp_path: Path) -> None:
    path = _autofill_workbook(tmp_path)
    graph = create_dependency_graph(
        path,
        ["Sheet1!B1", "Sheet1!B2", "Sheet1!C1"],
        load_values=True,
        warm_formula_shapes=True,
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["Sheet1!B1"])["Sheet1!B1"] == 20.0
        assert ev.evaluate(["Sheet1!B2"])["Sheet1!B2"] == 40.0
        assert ev.evaluate(["Sheet1!C1"])["Sheet1!C1"] == 20.0


def test_json_cache_round_trips_relative_axes(tmp_path: Path) -> None:
    from excel_grapher.grapher.cache import (
        GRAPH_CACHE_SCHEMA_VERSION,
        dependency_graph_from_json,
        dependency_graph_to_json,
    )

    path = _autofill_workbook(tmp_path)
    graph = create_dependency_graph(path, ["Sheet1!B2"], load_values=True)
    assert GRAPH_CACHE_SCHEMA_VERSION >= 8
    restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    original = graph.get_node("Sheet1!B2")
    loaded = restored.get_node("Sheet1!B2")
    assert original is not None and loaded is not None
    assert loaded.formula_ast == original.formula_ast
    assert loaded.normalized_formula == original.normalized_formula
    with FormulaEvaluator(restored) as ev:
        assert ev.evaluate(["Sheet1!B2"])["Sheet1!B2"] == 40.0
