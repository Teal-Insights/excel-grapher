"""Unit tests for compression placeholder nodes in core formula AST."""

from __future__ import annotations

import pytest

from excel_grapher.core.formula_ast import ColumnVarCellRefNode, SubexpressionRefNode


def test_column_var_cell_ref_defaults() -> None:
    node = ColumnVarCellRefNode()
    assert node.column_variable == "COL"
    assert node.sheet is None
    assert node.row is None


def test_column_var_cell_ref_with_sheet_and_row() -> None:
    node = ColumnVarCellRefNode(column_variable="COL", sheet="Ext", row=87)
    assert node == ColumnVarCellRefNode(column_variable="COL", sheet="Ext", row=87)
    assert node != ColumnVarCellRefNode(column_variable="ROW", sheet="Ext", row=87)


def test_column_var_cell_ref_rejects_empty_variable() -> None:
    with pytest.raises(ValueError, match="column_variable"):
        ColumnVarCellRefNode(column_variable="")


def test_subexpression_ref_accepts_cse_keys() -> None:
    node = SubexpressionRefNode(ref_key="_cse!0")
    assert node == SubexpressionRefNode(ref_key="_cse!0")
    assert SubexpressionRefNode(ref_key="_cse!42").ref_key == "_cse!42"


def test_subexpression_ref_rejects_invalid_keys() -> None:
    with pytest.raises(ValueError, match="ref_key"):
        SubexpressionRefNode(ref_key="cse!0")
    with pytest.raises(ValueError, match="ref_key"):
        SubexpressionRefNode(ref_key="_cse!")
    with pytest.raises(ValueError, match="ref_key"):
        SubexpressionRefNode(ref_key="")
