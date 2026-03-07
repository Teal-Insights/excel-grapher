from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.expr_eval import Unsupported, evaluate_expr
from excel_grapher.core.formula_ast import (
    BoolNode,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    parse,
)
from excel_grapher.core.types import CellValue


def _make_get_cell_value(values: Mapping[str, CellValue]):
    def get_cell_value(address: str) -> CellValue:
        return values[address]

    return get_cell_value


def test_core_formula_ast_parses_basic_literals_and_refs() -> None:
    assert parse("=1") == NumberNode(1.0)
    assert parse('="x"') == StringNode("x")
    assert parse("=TRUE") == BoolNode(True)
    assert parse("=Sheet1!A1") == CellRefNode("Sheet1!A1")
    assert parse("=Sheet1!A1:B2") == RangeNode("Sheet1!A1", "Sheet1!B2")

    ast = parse("=SUM(Sheet1!A1:A3)")
    assert isinstance(ast, FunctionCallNode)
    assert ast.name == "SUM"
    assert ast.args == [RangeNode("Sheet1!A1", "Sheet1!A3")]


def test_core_expr_eval_basic_functions_over_integers() -> None:
    values: dict[str, CellValue] = {
        "Sheet1!A1": 1,
        "Sheet1!A2": 2,
        "Sheet1!A3": 3,
        "Sheet1!B1": -5,
    }
    get_cell_value = _make_get_cell_value(values)

    # SUM over a 1D range.
    ast = parse("=SUM(Sheet1!A1:Sheet1!A3)")
    assert evaluate_expr(ast, get_cell_value=get_cell_value) == 6.0

    # MIN/MAX over the same range.
    ast = parse("=MIN(Sheet1!A1:Sheet1!A3)")
    assert evaluate_expr(ast, get_cell_value=get_cell_value) == 1.0

    ast = parse("=MAX(Sheet1!A1:Sheet1!A3)")
    assert evaluate_expr(ast, get_cell_value=get_cell_value) == 3.0

    # ABS over a single cell reference.
    ast = parse("=ABS(Sheet1!B1)")
    assert evaluate_expr(ast, get_cell_value=get_cell_value) == 5.0

    # IF over simple boolean conditions.
    ast = parse("=IF(TRUE, 1, 2)")
    assert evaluate_expr(ast, get_cell_value=get_cell_value) == 1.0

    ast = parse("=IF(FALSE, 1, 2)")
    assert evaluate_expr(ast, get_cell_value=get_cell_value) == 2.0


def test_core_expr_eval_unsupported_function_returns_sentinel() -> None:
    ast = parse("=FOO(1, 2)")
    result = evaluate_expr(ast, get_cell_value=lambda addr: 0)
    assert isinstance(result, Unsupported)


def test_core_expr_eval_respects_max_depth() -> None:
    # Build a modestly nested IF expression.
    # =IF(TRUE, IF(TRUE, IF(TRUE, 1, 0), 0), 0)
    formula = "=IF(TRUE, IF(TRUE, IF(TRUE, 1, 0), 0), 0)"
    ast = parse(formula)

    # With a very small max_depth, evaluation should give an Unsupported sentinel.
    shallow_result = evaluate_expr(
        ast, get_cell_value=lambda addr: 0, max_depth=1
    )
    assert isinstance(shallow_result, Unsupported)

    # With a generous max_depth, evaluation should succeed.
    deep_result = evaluate_expr(
        ast, get_cell_value=lambda addr: 0, max_depth=10
    )
    assert deep_result == 1.0


def test_core_expr_eval_row_and_column() -> None:
    # ROW(ref) returns the row number of the reference (value in ref is ignored).
    ast = parse("=ROW(Sheet1!B106)")
    assert evaluate_expr(ast, get_cell_value=lambda addr: 0) == 106

    ast = parse("=COLUMN(Sheet1!B106)")
    assert evaluate_expr(ast, get_cell_value=lambda addr: 0) == 2

    # ROW() and COLUMN() with no args require context.
    ast = parse("=ROW()")
    result = evaluate_expr(ast, get_cell_value=lambda addr: 0)
    assert isinstance(result, Unsupported)

    ast = parse("=ROW()")
    assert evaluate_expr(
        ast, get_cell_value=lambda addr: 0, context={"row": 106, "column": 1}
    ) == 106

    ast = parse("=ROW()-ROW(Sheet1!B106)+1")
    assert evaluate_expr(
        ast, get_cell_value=lambda addr: 0, context={"row": 106, "column": 1}
    ) == 1

