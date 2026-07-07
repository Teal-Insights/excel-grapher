"""Unit tests for compression parity harness."""

from __future__ import annotations

from excel_grapher.compression.nodes import ParallelFormulaNode
from excel_grapher.compression.parity import (
    assert_compression_parity,
    compare_compression_parity,
    compression_values_equal,
)
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    ColumnVarCellRefNode,
    FunctionCallNode,
    NumberNode,
    SubexpressionRefNode,
)
from excel_grapher.core.types import XlError

from .conftest import parse_formula


def test_compression_values_equal_accepts_float_tolerance() -> None:
    assert compression_values_equal(0.1 + 0.2, 0.3, rtol=1e-9, atol=1e-9)
    assert not compression_values_equal(0.1 + 0.2, 0.31, rtol=1e-9, atol=1e-9)


def test_compression_values_equal_matches_xl_error_codes() -> None:
    assert compression_values_equal(XlError.NA, XlError.NA)
    assert not compression_values_equal(XlError.NA, XlError.DIV)


def test_parity_numeric_identity_compression() -> None:
    input_values = {"Sheet1!A1": 10}
    original = {"Sheet1!B1": parse_formula("=Sheet1!A1+1")}
    compressed = {"Sheet1!B1": original["Sheet1!B1"]}
    assert_compression_parity(original, compressed, input_values=input_values)


def test_parity_reports_value_mismatch() -> None:
    input_values = {"Sheet1!A1": 10}
    original = {"Sheet1!B1": parse_formula("=Sheet1!A1+1")}
    compressed = {"Sheet1!B1": parse_formula("=Sheet1!A1+2")}
    mismatches = compare_compression_parity(
        original,
        compressed,
        input_values=input_values,
    )
    assert len(mismatches) == 1
    assert mismatches[0].cell_key == "Sheet1!B1"
    assert mismatches[0].original_value == 11
    assert mismatches[0].expanded_value == 12


def test_parity_error_codes() -> None:
    input_values = {"Sheet1!A1": 0}
    original = {"Sheet1!B1": parse_formula("=Sheet1!A1/Sheet1!A1")}
    compressed = {"Sheet1!B1": original["Sheet1!B1"]}
    assert_compression_parity(original, compressed, input_values=input_values)


def test_parity_na_formula() -> None:
    original = {"Sheet1!A1": parse_formula("=NA()")}
    compressed = {"Sheet1!A1": original["Sheet1!A1"]}
    assert_compression_parity(original, compressed, input_values={})


def test_parity_parallel_row_fixture() -> None:
    input_values = {"Ext!D3": "Yes"}
    original = {
        "'Chart Data'!D177": FunctionCallNode(
            "IF",
            [
                CellRefNode("Ext!D3"),
                parse_formula("=NA()"),
                CellRefNode("Ext!D87"),
            ],
        ),
        "'Chart Data'!E177": FunctionCallNode(
            "IF",
            [
                CellRefNode("Ext!D3"),
                parse_formula("=NA()"),
                CellRefNode("Ext!E87"),
            ],
        ),
    }
    compressed = {
        "parallel:Chart Data!177": ParallelFormulaNode(
            sheet="Chart Data",
            template=FunctionCallNode(
                "IF",
                [
                    CellRefNode("Ext!D3"),
                    parse_formula("=NA()"),
                    ColumnVarCellRefNode(sheet="Ext", row=87),
                ],
            ),
            start_col="D",
            end_col="E",
            output_row=177,
        ),
    }
    assert_compression_parity(original, compressed, input_values=input_values)


def test_parity_cse_fixture() -> None:
    input_values = {"Sheet1!B1": 2, "Sheet1!C1": 3}
    shared = parse_formula("=Sheet1!B1+Sheet1!C1")
    original = {
        "Sheet1!A1": BinaryOpNode("*", shared, NumberNode(3.0)),
        "Sheet1!A2": BinaryOpNode("+", shared, NumberNode(10.0)),
    }
    compressed = {
        "_cse!0": shared,
        "Sheet1!A1": BinaryOpNode("*", SubexpressionRefNode("_cse!0"), NumberNode(3.0)),
        "Sheet1!A2": BinaryOpNode("+", SubexpressionRefNode("_cse!0"), NumberNode(10.0)),
    }
    assert_compression_parity(original, compressed, input_values=input_values)
