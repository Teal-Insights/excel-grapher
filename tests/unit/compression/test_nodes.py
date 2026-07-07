"""Unit tests for compression artifact node types."""

from __future__ import annotations

import pytest

from excel_grapher.compression.nodes import ParallelFormulaNode, TacoPatternNode
from excel_grapher.core.formula_ast import (
    CellRefNode,
    ColumnVarCellRefNode,
    FunctionCallNode,
    NumberNode,
)
from excel_grapher.grapher.range_compression.types import Orientation, PatternKind

from .conftest import parse_formula


def test_parallel_formula_node_template_primary() -> None:
    template = FunctionCallNode(
        "IF",
        [
            CellRefNode("Ext!$D$3"),
            parse_formula("=NA()"),
            ColumnVarCellRefNode(sheet="Ext", row=87),
        ],
    )
    node = ParallelFormulaNode(
        sheet="Chart Data",
        template=template,
        start_col="D",
        end_col="X",
        output_row=177,
    )
    assert node.sheet == "Chart Data"
    assert node.start_col == "D"
    assert node.end_col == "X"
    assert node.output_row == 177
    assert node.column_variable == "COL"
    assert node.condition is None
    assert node.if_true is None
    assert node.if_false is None
    assert node.template is template


def test_parallel_formula_node_optional_if_projection() -> None:
    condition = CellRefNode("Ext!$D$3")
    if_true = parse_formula("=NA()")
    if_false = ColumnVarCellRefNode(sheet="Ext", row=87)
    node = ParallelFormulaNode(
        sheet="Chart Data",
        template=FunctionCallNode("IF", [condition, if_true, if_false]),
        start_col="D",
        end_col="F",
        output_row=177,
        condition=condition,
        if_true=if_true,
        if_false=if_false,
    )
    assert node.condition is condition
    assert node.if_true is if_true
    assert node.if_false is if_false


def test_parallel_formula_node_rejects_invalid_column_range() -> None:
    template = NumberNode(1.0)
    with pytest.raises(ValueError, match="start_col"):
        ParallelFormulaNode(
            sheet="Sheet1",
            template=template,
            start_col="Z",
            end_col="A",
            output_row=1,
        )


def test_parallel_formula_node_rejects_invalid_output_row() -> None:
    template = NumberNode(1.0)
    with pytest.raises(ValueError, match="output_row"):
        ParallelFormulaNode(
            sheet="Sheet1",
            template=template,
            start_col="A",
            end_col="C",
            output_row=0,
        )


def test_parallel_formula_node_column_count() -> None:
    template = NumberNode(1.0)
    node = ParallelFormulaNode(
        sheet="Sheet1",
        template=template,
        start_col="D",
        end_col="F",
        output_row=10,
    )
    assert node.column_count == 3


def test_taco_pattern_node_construction() -> None:
    template = parse_formula("=Sheet1!B2*Sheet1!C2")
    node = TacoPatternNode(
        kind=PatternKind.rr,
        sheet="Sheet1",
        min_col="B",
        min_row=3,
        max_col="B",
        max_row=7,
        template=template,
        orientation=Orientation.column,
    )
    assert node.kind is PatternKind.rr
    assert node.orientation is Orientation.column
    assert node.template is template


def test_taco_pattern_node_rejects_invalid_bounds() -> None:
    template = parse_formula("=Sheet1!B2")
    with pytest.raises(ValueError, match="min_col"):
        TacoPatternNode(
            kind=PatternKind.rr,
            sheet="Sheet1",
            min_col="C",
            min_row=3,
            max_col="B",
            max_row=7,
            template=template,
            orientation=Orientation.column,
        )
    with pytest.raises(ValueError, match="min_row"):
        TacoPatternNode(
            kind=PatternKind.rr,
            sheet="Sheet1",
            min_col="B",
            min_row=8,
            max_col="B",
            max_row=3,
            template=template,
            orientation=Orientation.column,
        )


def test_taco_pattern_node_rejects_single_kind() -> None:
    template = parse_formula("=Sheet1!B2")
    with pytest.raises(ValueError, match="single"):
        TacoPatternNode(
            kind=PatternKind.single,
            sheet="Sheet1",
            min_col="B",
            min_row=3,
            max_col="B",
            max_row=7,
            template=template,
            orientation=Orientation.column,
        )


def test_taco_pattern_node_cell_count() -> None:
    template = parse_formula("=Sheet1!B2*Sheet1!C2")
    node = TacoPatternNode(
        kind=PatternKind.rr,
        sheet="Sheet1",
        min_col="B",
        min_row=3,
        max_col="C",
        max_row=5,
        template=template,
        orientation=Orientation.column,
    )
    assert node.cell_count == 6
