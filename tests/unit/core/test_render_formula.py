"""`render_formula` styles and `coerce_relative_refs` (#543)."""

from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    CellRef,
    CellRefNode,
    FormulaStyle,
    RangeNode,
    RelativeAxis,
    WholeColumnNode,
    WholeRowNode,
    parse,
    parse_preserving_axes,
    render_formula,
    unparse_normalized_formula,
)


def test_a1_absolute_matches_unparse_normalized_formula() -> None:
    ast = parse_preserving_axes("=A1*2", anchor="Sheet1!B1")
    rendered = render_formula(ast, anchor="Sheet1!B1", style=FormulaStyle.A1_ABSOLUTE)
    assert rendered == "=Sheet1!A1*2"
    assert rendered == unparse_normalized_formula(ast, anchor="Sheet1!B1")


def test_a1_excel_preserves_dollar_markers_and_omits_same_sheet() -> None:
    ast = parse_preserving_axes("=$A$1+$A1+A$1", anchor="Sheet1!B2")
    assert render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.A1_EXCEL) == "=$A$1+$A1+A$1"


def test_a1_excel_qualifies_cross_sheet_refs() -> None:
    ast = parse_preserving_axes("=Other!B5+1", anchor="Sheet1!A1")
    assert render_formula(ast, anchor="Sheet1!A1", style=FormulaStyle.A1_EXCEL) == "=Other!B5+1"


def test_r1c1_relative_and_absolute_axes() -> None:
    ast = parse_preserving_axes("=$A$1+$A1+A$1", anchor="Sheet1!B2")
    assert (
        render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.R1C1) == "=R1C1+R[-1]C1+R1C[-1]"
    )


def test_r1c1_zero_offset_omits_brackets() -> None:
    ast = parse_preserving_axes("=A2+B1", anchor="Sheet1!B2")
    assert render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.R1C1) == "=RC[-1]+R[-1]C"


def test_r1c1_qualifies_cross_sheet_refs() -> None:
    ast = parse_preserving_axes("=Other!A1", anchor="Sheet1!B2")
    assert render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.R1C1) == "=Other!R[-1]C[-1]"


def test_coerce_relative_refs_absolutizes_a1_excel() -> None:
    ast = parse_preserving_axes("=A1+$B$1", anchor="Sheet1!B2")
    assert render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.A1_EXCEL) == "=A1+$B$1"
    assert (
        render_formula(
            ast,
            anchor="Sheet1!B2",
            style=FormulaStyle.A1_EXCEL,
            coerce_relative_refs=True,
        )
        == "=$A$1+$B$1"
    )


def test_coerce_relative_refs_absolutizes_r1c1() -> None:
    ast = parse_preserving_axes("=A1+$B$1", anchor="Sheet1!B2")
    assert (
        render_formula(
            ast,
            anchor="Sheet1!B2",
            style=FormulaStyle.R1C1,
            coerce_relative_refs=True,
        )
        == "=R1C1+R1C2"
    )


def test_coerce_relative_refs_is_noop_for_a1_absolute() -> None:
    ast = parse_preserving_axes("=A1*2", anchor="Sheet1!B1")
    assert render_formula(
        ast,
        anchor="Sheet1!B1",
        style=FormulaStyle.A1_ABSOLUTE,
        coerce_relative_refs=True,
    ) == render_formula(ast, anchor="Sheet1!B1", style=FormulaStyle.A1_ABSOLUTE)


def test_relative_axes_require_anchor() -> None:
    ast = parse_preserving_axes("=A1", anchor="Sheet1!B2")
    with pytest.raises(ValueError, match="anchor"):
        render_formula(ast, style=FormulaStyle.A1_ABSOLUTE)
    with pytest.raises(ValueError, match="anchor"):
        render_formula(ast, style=FormulaStyle.A1_EXCEL)
    with pytest.raises(ValueError, match="anchor"):
        render_formula(
            ast,
            style=FormulaStyle.R1C1,
            coerce_relative_refs=True,
        )


def test_render_range_styles() -> None:
    ast = parse_preserving_axes("=SUM(A$1:B2)", anchor="Sheet1!B2")
    assert (
        render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.A1_ABSOLUTE)
        == "=SUM(Sheet1!A1:B2)"
    )
    assert render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.A1_EXCEL) == "=SUM(A$1:B2)"
    assert render_formula(ast, anchor="Sheet1!B2", style=FormulaStyle.R1C1) == "=SUM(R1C[-1]:RC)"


def test_render_whole_column_and_row() -> None:
    col = WholeColumnNode(sheet="Data", col=RelativeAxis(-1))
    row = WholeRowNode(sheet="Data", row=AbsoluteAxis(4))
    assert render_formula(col, anchor="Data!B1", style=FormulaStyle.A1_ABSOLUTE) == "=Data!A:A"
    assert render_formula(col, anchor="Data!B1", style=FormulaStyle.A1_EXCEL) == "=A:A"
    assert render_formula(col, anchor="Data!B1", style=FormulaStyle.R1C1) == "=C[-1]:C[-1]"
    assert render_formula(row, anchor="Sheet1!A1", style=FormulaStyle.A1_EXCEL) == "=Data!$4:$4"
    assert render_formula(row, anchor="Sheet1!A1", style=FormulaStyle.R1C1) == "=Data!R4:R4"


def test_render_quoted_sheet_name() -> None:
    ast = parse("='Other Sheet'!B5+1")
    assert render_formula(ast, style=FormulaStyle.A1_ABSOLUTE) == "='Other Sheet'!B5+1"
    assert (
        render_formula(ast, anchor="Sheet1!A1", style=FormulaStyle.A1_EXCEL)
        == "='Other Sheet'!$B$5+1"
    )
    assert (
        render_formula(ast, anchor="Sheet1!A1", style=FormulaStyle.R1C1) == "='Other Sheet'!R5C2+1"
    )


def test_style_accepts_string_value() -> None:
    ast = CellRefNode(CellRef(sheet="Sheet1", col=RelativeAxis(-1), row=RelativeAxis(0)))
    assert render_formula(ast, anchor="Sheet1!B1", style="r1c1") == "=RC[-1]"


def test_fully_absolute_cell_ref_node_r1c1() -> None:
    ast = CellRefNode(
        CellRef(sheet="Sheet1", col=AbsoluteAxis(1), row=AbsoluteAxis(1)),
    )
    assert render_formula(ast, style=FormulaStyle.R1C1) == "=Sheet1!R1C1"
    mixed = RangeNode(
        CellRef(sheet="Sheet1", col=AbsoluteAxis(1), row=RelativeAxis(-1)),
        CellRef(sheet="Sheet1", col=RelativeAxis(0), row=AbsoluteAxis(3)),
    )
    assert render_formula(mixed, anchor="Sheet1!B2", style=FormulaStyle.R1C1) == "=R[-1]C1:R3C"
