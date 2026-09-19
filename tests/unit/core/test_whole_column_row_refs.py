from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    CellRefNode,
    FunctionCallNode,
    RangeNode,
    WholeColumnNode,
    WholeRowNode,
    parse,
    unparse_normalized_formula,
)
from excel_grapher.core.formula_normalization import (
    expand_whole_column_row_for_parse,
    normalize_excel_formula,
)
from excel_grapher.core.range_shorthand import (
    EXCEL_MAX_ROW,
    expand_whole_column_deps,
    expand_whole_column_span_deps,
    expand_whole_row_span_deps,
    resolve_whole_column,
    resolve_whole_column_span,
    resolve_whole_row,
    resolve_whole_row_span,
)
from excel_grapher.core.types import ExcelRange
from excel_grapher.grapher.parser import parse_range_refs_with_spans


@pytest.mark.parametrize(
    "formula",
    [
        "='QB - Stafford'!C:C",
        "='QB - Stafford'!$C:$C",
        "=Data!A:A",
        "=Data!5:5",
        "='My Sheet'!B:B",
    ],
)
def test_parse_whole_column_or_row_shorthand(formula: str) -> None:
    ast = parse(formula)
    assert isinstance(ast, (WholeColumnNode, WholeRowNode))


def test_parse_whole_column_ast_shape() -> None:
    assert parse("=Data!C:C") == WholeColumnNode(sheet="Data", column="C")


def test_parse_whole_row_ast_shape() -> None:
    assert parse("=Data!5:5") == WholeRowNode(sheet="Data", row=5)


def test_parse_whole_column_span_is_open_axis_node() -> None:
    node = parse("=Data!A:C")
    assert isinstance(node, WholeColumnNode)
    assert not isinstance(node, RangeNode)
    assert node == WholeColumnNode(sheet="Data", start_col="A", end_col="C")
    assert node.start_col == AbsoluteAxis(1)
    assert node.end_col == AbsoluteAxis(3)


def test_parse_whole_row_span_is_open_axis_node() -> None:
    node = parse("=Data!5:10")
    assert isinstance(node, WholeRowNode)
    assert not isinstance(node, RangeNode)
    assert node == WholeRowNode(sheet="Data", start_row=5, end_row=10)


def test_parse_sum_over_whole_column_span() -> None:
    ast = parse("=SUM(Data!A:C)")
    assert isinstance(ast, FunctionCallNode)
    assert ast.name == "SUM"
    assert ast.args[0] == WholeColumnNode(sheet="Data", start_col="A", end_col="C")


def test_parse_preserves_reversed_whole_axis_spans() -> None:
    assert parse("=Data!C:A") == WholeColumnNode(sheet="Data", start_col="C", end_col="A")
    assert parse("=Data!10:5") == WholeRowNode(sheet="Data", start_row=10, end_row=5)


def test_unparse_whole_axis_spans_round_trip() -> None:
    for formula in ("=Data!A:C", "=Data!5:10", "=SUM(Data!A:C)", "=Data!C:A"):
        ast = parse(formula)
        rendered = unparse_normalized_formula(ast)
        assert rendered == formula
        assert parse(rendered) == ast
        assert "A1:" not in rendered
        assert "C1048576" not in rendered


def test_parse_index_match_whole_column_formula() -> None:
    formula = (
        "=INDEX(Data!C:C,"
        "MATCH(INDEX(Stats!C1:G1,MATCH(MAX(Stats!C2:G2),Stats!C2:G2,0)),"
        "Data!A:A,0))"
    )
    ast = parse(formula)
    assert isinstance(ast, FunctionCallNode)
    assert ast.name == "INDEX"
    assert isinstance(ast.args[0], WholeColumnNode)


def test_parse_whole_column_does_not_break_cell_ref() -> None:
    assert parse("='QB - Stafford'!C1") == CellRefNode("'QB - Stafford'!C1")


def test_resolve_whole_column_uses_workbook_bounds() -> None:
    bounds = {"Data": (50, 10)}
    rng = resolve_whole_column("Data", "C", bounds)
    assert rng == ExcelRange("Data", 1, 3, 50, 3)


def test_resolve_whole_column_quoted_sheet_name() -> None:
    bounds = {"QB - Stafford": (30, 5)}
    rng = resolve_whole_column("QB - Stafford", "A", bounds)
    assert rng.end_row == 30


def test_resolve_whole_row_uses_workbook_bounds() -> None:
    bounds = {"Data": (50, 10)}
    rng = resolve_whole_row("Data", 5, bounds)
    assert rng.start_col == 1 and rng.end_col == 10 and rng.start_row == rng.end_row == 5


def test_resolve_whole_column_defaults_to_excel_max_without_bounds() -> None:
    rng = resolve_whole_column("Missing", "A", {})
    assert rng.end_row == EXCEL_MAX_ROW


def test_resolve_whole_row_span_uses_used_columns() -> None:
    bounds = {"Macrofw": (6, 2)}
    rng = resolve_whole_row_span("Macrofw", 5, 10, bounds)
    assert rng == ExcelRange("Macrofw", 5, 1, 10, 2)


def test_resolve_whole_row_span_normalizes_reversed_rows() -> None:
    bounds = {"Data": (20, 4)}
    rng = resolve_whole_row_span("Data", 10, 5, bounds)
    assert rng == ExcelRange("Data", 5, 1, 10, 4)


def test_resolve_whole_column_span_uses_used_rows() -> None:
    bounds = {"Data": (20, 8)}
    rng = resolve_whole_column_span("Data", "B", "D", bounds)
    assert rng == ExcelRange("Data", 1, 2, 20, 4)


def test_resolve_whole_column_span_normalizes_reversed_cols() -> None:
    bounds = {"Data": (10, 5)}
    rng = resolve_whole_column_span("Data", "D", "B", bounds)
    assert rng == ExcelRange("Data", 1, 2, 10, 4)


def test_expand_whole_column_deps_enumerates_used_range() -> None:
    bounds = {"Data": (3, 2)}
    deps = expand_whole_column_deps("Data", "A", bounds)
    assert deps == [("Data", "A1"), ("Data", "A2"), ("Data", "A3")]


def test_expand_whole_column_span_deps_uses_used_rows() -> None:
    bounds = {"Data": (2, 8)}
    deps = expand_whole_column_span_deps("Data", "A", "C", bounds)
    assert deps == [
        ("Data", "A1"),
        ("Data", "B1"),
        ("Data", "C1"),
        ("Data", "A2"),
        ("Data", "B2"),
        ("Data", "C2"),
    ]


def test_expand_whole_row_span_deps_uses_used_columns() -> None:
    bounds = {"Data": (20, 2)}
    deps = expand_whole_row_span_deps("Data", 5, 6, bounds)
    assert deps == [
        ("Data", "A5"),
        ("Data", "B5"),
        ("Data", "A6"),
        ("Data", "B6"),
    ]


def test_normalize_preserves_whole_column_shorthand() -> None:
    out = normalize_excel_formula("=MATCH(x,Data!$A:$A,0)", "Sheet1")
    assert out == "=MATCH(x,Data!A:A,0)"


def test_normalize_qualifies_local_whole_column() -> None:
    out = normalize_excel_formula("=MATCH(x,A:A,0)", "Sheet1")
    assert out == "=MATCH(x,Sheet1!A:A,0)"


def test_normalize_preserves_whole_axis_spans_without_expanding() -> None:
    assert normalize_excel_formula("=SUM(Data!$A:$C)", "Sheet1") == "=SUM(Data!A:C)"
    assert normalize_excel_formula("=SUM(Data!$5:$10)", "Sheet1") == "=SUM(Data!5:10)"
    assert normalize_excel_formula("=SUM(A:C)", "Sheet1") == "=SUM(Sheet1!A:C)"
    assert normalize_excel_formula("=SUM(5:10)", "Sheet1") == "=SUM(Sheet1!5:10)"


def test_expand_whole_column_row_for_parse_quoted_sheet() -> None:
    bounds = {"QB - Stafford": (30, 5)}
    formula = "=INDEX('QB - Stafford'!C:C,1)"
    expanded = expand_whole_column_row_for_parse(formula, bounds)
    assert "'QB - Stafford'!C1:'QB - Stafford'!C30" in expanded


def test_parse_range_refs_whole_column_quoted() -> None:
    refs = parse_range_refs_with_spans("=INDEX('Data'!C:C,1)")
    assert len(refs) == 1
    start, end, _span = refs[0]
    assert start.sheet == "Data"
    assert start.column == end.column == "C"
    assert start.range_kind == "whole_column"


def test_parse_range_refs_whole_row() -> None:
    refs = parse_range_refs_with_spans("=SUM('Data'!5:5)")
    assert len(refs) == 1
    start, end, _span = refs[0]
    assert start.sheet == "Data"
    assert start.row == end.row == 5
    assert start.range_kind == "whole_row"


def test_parse_range_refs_with_spans_whole_column() -> None:
    refs = parse_range_refs_with_spans("=MATCH(x,'Data'!A:A,0)")
    assert len(refs) == 1
    _, _, span = refs[0]
    assert span == (9, 19)


def test_parse_range_refs_whole_column_span() -> None:
    refs = parse_range_refs_with_spans("=SUM(Data!A:C)")
    assert len(refs) == 1
    start, end, _span = refs[0]
    assert start.sheet == end.sheet == "Data"
    assert start.column == "A"
    assert end.column == "C"
    assert start.range_kind == end.range_kind == "whole_column"


def test_parse_range_refs_whole_row_span() -> None:
    refs = parse_range_refs_with_spans("=SUM(Data!5:10)")
    assert len(refs) == 1
    start, end, _span = refs[0]
    assert start.sheet == end.sheet == "Data"
    assert start.row == 5
    assert end.row == 10
    assert start.range_kind == end.range_kind == "whole_row"
