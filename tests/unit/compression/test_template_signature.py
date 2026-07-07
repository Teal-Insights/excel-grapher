"""Unit tests for parallel-row template signatures and column normalization."""

from __future__ import annotations

from excel_grapher.compression.template_signature import (
    collect_cell_ref_addresses,
    fixed_cell_refs_in_group,
    template_signature,
    with_column_variable,
)
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    ColumnVarCellRefNode,
    FunctionCallNode,
    NumberNode,
)

from .conftest import parse_formula


def _if_row_asts() -> tuple:
    """LIC-DSF-style IF row (Chart Data row 177 pattern)."""
    d = parse_formula('=IF(Ext!D3="No",NA(),Ext!D87)')
    e = parse_formula('=IF(Ext!D3="No",NA(),Ext!E87)')
    f = parse_formula('=IF(Ext!D3="No",NA(),Ext!F87)')
    return d, e, f


def _multiply_row_asts() -> tuple:
    """Shared prefix/suffix without top-level IF."""
    d = parse_formula("=Sheet1!D10*2")
    e = parse_formula("=Sheet1!E10*2")
    f = parse_formula("=Sheet1!F10*2")
    return d, e, f


def test_collect_cell_ref_addresses() -> None:
    d, e, _ = _if_row_asts()
    assert collect_cell_ref_addresses(d) == {"Ext!D3", "Ext!D87"}
    assert collect_cell_ref_addresses(e) == {"Ext!D3", "Ext!E87"}


def test_fixed_cell_refs_in_group_if_row() -> None:
    d, e, f = _if_row_asts()
    assert fixed_cell_refs_in_group((d, e, f)) == frozenset({"Ext!D3"})


def test_fixed_cell_refs_in_group_multiply_row() -> None:
    d, e, f = _multiply_row_asts()
    assert fixed_cell_refs_in_group((d, e, f)) == frozenset()


def test_with_column_variable_if_row() -> None:
    d, e, f = _if_row_asts()
    peers = (d, e, f)
    norm_d = with_column_variable(
        d,
        output_sheet="Chart Data",
        output_col="D",
        output_row=177,
        peer_asts=peers,
    )
    norm_e = with_column_variable(
        e,
        output_sheet="Chart Data",
        output_col="E",
        output_row=177,
        peer_asts=peers,
    )
    assert isinstance(norm_d, FunctionCallNode)
    assert norm_d.args[0] == d.args[0]
    assert norm_d.args[1] == d.args[1]
    assert norm_d.args[2] == ColumnVarCellRefNode(column_variable="COL", sheet="Ext", row=87)
    assert template_signature(norm_d) == template_signature(norm_e)


def test_with_column_variable_multiply_row() -> None:
    d, e, f = _multiply_row_asts()
    peers = (d, e, f)
    norm_d = with_column_variable(
        d,
        output_sheet="Sheet1",
        output_col="D",
        output_row=1,
        peer_asts=peers,
    )
    norm_e = with_column_variable(
        e,
        output_sheet="Sheet1",
        output_col="E",
        output_row=1,
        peer_asts=peers,
    )
    assert template_signature(norm_d) == template_signature(norm_e)
    assert norm_d == BinaryOpNode(
        "*",
        ColumnVarCellRefNode(column_variable="COL", sheet="Sheet1", row=10),
        NumberNode(2.0),
    )


def test_with_column_variable_keeps_fixed_refs() -> None:
    d, e, f = _if_row_asts()
    norm_d = with_column_variable(
        d,
        output_sheet="Chart Data",
        output_col="D",
        output_row=177,
        peer_asts=(d, e, f),
    )
    assert isinstance(norm_d, FunctionCallNode)
    condition = norm_d.args[0]
    assert collect_cell_ref_addresses(condition) == {"Ext!D3"}


def test_template_signature_matches_across_columns() -> None:
    d, e, f = _multiply_row_asts()
    peers = (d, e, f)
    signatures = {
        template_signature(
            with_column_variable(
                ast,
                output_sheet="Sheet1",
                output_col=col,
                output_row=1,
                peer_asts=peers,
            )
        )
        for ast, col in zip((d, e, f), ("D", "E", "F"), strict=True)
    }
    assert len(signatures) == 1


def test_template_signature_rejects_different_operands() -> None:
    left = parse_formula("=Sheet1!D10*2")
    right = parse_formula("=Sheet1!D10*3")
    peers_left = (left, parse_formula("=Sheet1!E10*2"), parse_formula("=Sheet1!F10*2"))
    peers_right = (right, parse_formula("=Sheet1!E10*3"), parse_formula("=Sheet1!F10*3"))
    sig_left = template_signature(
        with_column_variable(
            left,
            output_sheet="Sheet1",
            output_col="D",
            output_row=1,
            peer_asts=peers_left,
        )
    )
    sig_right = template_signature(
        with_column_variable(
            right,
            output_sheet="Sheet1",
            output_col="D",
            output_row=1,
            peer_asts=peers_right,
        )
    )
    assert sig_left != sig_right


def test_template_signature_rejects_different_functions() -> None:
    if_row = parse_formula('=IF(Ext!D3="No",NA(),Ext!D87)')
    sum_row = parse_formula("=SUM(Ext!D87)")
    assert template_signature(if_row) != template_signature(sum_row)
