from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import (
    CellRefNode,
    FunctionCallNode,
    WholeColumnNode,
    WholeRowNode,
    parse,
)
from excel_grapher.core.formula_normalization import (
    expand_whole_column_row_for_parse,
    normalize_excel_formula,
)
from excel_grapher.core.range_shorthand import (
    EXCEL_MAX_ROW,
    expand_whole_column_deps,
    resolve_whole_column,
    resolve_whole_row,
)
from excel_grapher.core.types import ExcelRange
from excel_grapher.grapher.parser import parse_range_refs, parse_range_refs_with_spans


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


def test_expand_whole_column_deps_enumerates_used_range() -> None:
    bounds = {"Data": (3, 2)}
    deps = expand_whole_column_deps("Data", "A", bounds)
    assert deps == [("Data", "A1"), ("Data", "A2"), ("Data", "A3")]


def test_normalize_preserves_whole_column_shorthand() -> None:
    out = normalize_excel_formula("=MATCH(x,Data!$A:$A,0)", "Sheet1")
    assert out == "=MATCH(x,Data!A:A,0)"


def test_normalize_qualifies_local_whole_column() -> None:
    out = normalize_excel_formula("=MATCH(x,A:A,0)", "Sheet1")
    assert out == "=MATCH(x,Sheet1!A:A,0)"


def test_expand_whole_column_row_for_parse_quoted_sheet() -> None:
    bounds = {"QB - Stafford": (30, 5)}
    formula = "=INDEX('QB - Stafford'!C:C,1)"
    expanded = expand_whole_column_row_for_parse(formula, bounds)
    assert "'QB - Stafford'!C1:'QB - Stafford'!C30" in expanded


def test_parse_range_refs_whole_column_quoted() -> None:
    refs = parse_range_refs("=INDEX('Data'!C:C,1)")
    assert len(refs) == 1
    start, end = refs[0]
    assert start.sheet == "Data"
    assert start.column == end.column == "C"
    assert start.range_kind == "whole_column"


def test_parse_range_refs_whole_row() -> None:
    refs = parse_range_refs("=SUM('Data'!5:5)")
    assert len(refs) == 1
    start, end = refs[0]
    assert start.sheet == "Data"
    assert start.row == end.row == 5
    assert start.range_kind == "whole_row"


def test_parse_range_refs_with_spans_whole_column() -> None:
    refs = parse_range_refs_with_spans("=MATCH(x,'Data'!A:A,0)")
    assert len(refs) == 1
    _, _, span = refs[0]
    assert span == (9, 19)
