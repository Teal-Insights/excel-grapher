"""Unit tests for expand_compressed_to_cells."""

from __future__ import annotations

from excel_grapher.compression.expand import expand_compressed_to_cells
from excel_grapher.compression.nodes import ParallelFormulaNode, TacoPatternNode
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    ColumnVarCellRefNode,
    FunctionCallNode,
    NumberNode,
    SubexpressionRefNode,
)
from excel_grapher.grapher.range_compression.types import Orientation, PatternKind

from .conftest import parse_formula


def test_expand_identity_passes_through_per_cell_asts() -> None:
    a1 = parse_formula("=Sheet1!A1+1")
    b1 = parse_formula("=Sheet1!B1*2")
    compressed = {"Sheet1!A1": a1, "Sheet1!B1": b1}
    assert expand_compressed_to_cells(compressed) == {
        "Sheet1!A1": a1,
        "Sheet1!B1": b1,
    }


def test_expand_parallel_if_row_materializes_columns() -> None:
    template = FunctionCallNode(
        "IF",
        [
            CellRefNode("Ext!D3"),
            parse_formula("=NA()"),
            ColumnVarCellRefNode(sheet="Ext", row=87),
        ],
    )
    compressed = {
        "parallel:Chart Data!177": ParallelFormulaNode(
            sheet="Chart Data",
            template=template,
            start_col="D",
            end_col="F",
            output_row=177,
        ),
    }
    expanded = expand_compressed_to_cells(compressed)
    assert set(expanded) == {
        "'Chart Data'!D177",
        "'Chart Data'!E177",
        "'Chart Data'!F177",
    }
    assert expanded["'Chart Data'!D177"] == FunctionCallNode(
        "IF",
        [
            CellRefNode("Ext!D3"),
            parse_formula("=NA()"),
            CellRefNode("Ext!D87"),
        ],
    )
    assert expanded["'Chart Data'!E177"] == FunctionCallNode(
        "IF",
        [
            CellRefNode("Ext!D3"),
            parse_formula("=NA()"),
            CellRefNode("Ext!E87"),
        ],
    )


def test_expand_inlines_cse_bindings() -> None:
    hoisted = parse_formula("=Sheet1!B1+Sheet1!C1")
    compressed = {
        "_cse!0": hoisted,
        "Sheet1!A1": BinaryOpNode("*", SubexpressionRefNode("_cse!0"), NumberNode(3.0)),
        "Sheet1!A2": BinaryOpNode("+", SubexpressionRefNode("_cse!0"), NumberNode(10.0)),
    }
    expanded = expand_compressed_to_cells(compressed)
    assert expanded == {
        "Sheet1!A1": BinaryOpNode("*", hoisted, NumberNode(3.0)),
        "Sheet1!A2": BinaryOpNode("+", hoisted, NumberNode(10.0)),
    }


def test_expand_taco_rr_column_fill_down() -> None:
    template = parse_formula("=Sheet1!B3*Sheet1!C3")
    compressed = {
        "taco:Sheet1!D3:D5": TacoPatternNode(
            kind=PatternKind.rr,
            sheet="Sheet1",
            min_col="D",
            min_row=3,
            max_col="D",
            max_row=5,
            template=template,
            orientation=Orientation.column,
        ),
    }
    expanded = expand_compressed_to_cells(compressed)
    assert set(expanded) == {"Sheet1!D3", "Sheet1!D4", "Sheet1!D5"}
    assert expanded["Sheet1!D3"] == parse_formula("=Sheet1!B3*Sheet1!C3")
    assert expanded["Sheet1!D4"] == parse_formula("=Sheet1!B4*Sheet1!C4")
    assert expanded["Sheet1!D5"] == parse_formula("=Sheet1!B5*Sheet1!C5")


def test_expand_mixed_parallel_cse_and_plain_cells() -> None:
    hoisted = parse_formula("=Sheet1!B1+Sheet1!C1")
    template = BinaryOpNode("*", SubexpressionRefNode("_cse!0"), NumberNode(2.0))
    compressed = {
        "_cse!0": hoisted,
        "Sheet1!Z1": parse_formula("=Sheet1!Z1"),
        "parallel:Sheet1!1": ParallelFormulaNode(
            sheet="Sheet1",
            template=template,
            start_col="D",
            end_col="E",
            output_row=1,
        ),
    }
    expanded = expand_compressed_to_cells(compressed)
    assert expanded["Sheet1!Z1"] == parse_formula("=Sheet1!Z1")
    assert expanded["Sheet1!D1"] == BinaryOpNode("*", hoisted, NumberNode(2.0))
    assert expanded["Sheet1!E1"] == BinaryOpNode("*", hoisted, NumberNode(2.0))
