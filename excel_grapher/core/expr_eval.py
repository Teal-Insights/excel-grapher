from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import cast

from openpyxl.utils.cell import coordinate_to_tuple

from .coercions import flatten, numeric_values, to_bool, to_number
from .formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
)
from .operators import (
    xl_add,
    xl_concat,
    xl_div,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_le,
    xl_lt,
    xl_mul,
    xl_ne,
    xl_neg,
    xl_percent,
    xl_pow,
    xl_sub,
)
from .types import CellValue, ExcelRange, XlError


@dataclass(frozen=True, slots=True)
class Unsupported:
    """Sentinel result for expressions this restricted evaluator cannot handle."""

    reason: str | None = None


_FunctionsMapping = Mapping[str, Callable[[list[CellValue]], CellValue]]


def evaluate_expr(
    node: AstNode,
    *,
    get_cell_value: Callable[[str], CellValue],
    functions: _FunctionsMapping | None = None,
    max_depth: int = 10,
) -> CellValue | XlError | Unsupported:
    """Evaluate a restricted Excel expression AST against a cell-value callback.

    This evaluator is intentionally small and representation-agnostic. It
    supports:
    - Literals (numbers, strings, booleans, Excel error literals)
    - Sheet-qualified cell refs (via get_cell_value)
    - Simple ranges (resolved to numpy arrays via ExcelRange.resolve)
    - A small function whitelist (SUM, MIN, MAX, ABS, IF)
    - Basic arithmetic/comparison/string operators
    """

    funcs: _FunctionsMapping = functions or {}
    return _eval(node, get_cell_value, funcs, max_depth, depth=0)


def _eval(
    node: AstNode,
    get_cell_value: Callable[[str], CellValue],
    functions: _FunctionsMapping,
    max_depth: int,
    *,
    depth: int,
) -> CellValue | XlError | Unsupported:
    if depth > max_depth:
        return Unsupported("max_depth exceeded")

    if isinstance(node, NumberNode):
        return node.value

    if isinstance(node, StringNode):
        return node.value

    if isinstance(node, BoolNode):
        return node.value

    if isinstance(node, ErrorNode):
        return node.error

    if isinstance(node, CellRefNode):
        return get_cell_value(node.address)

    if isinstance(node, RangeNode):
        excel_range = _range_node_to_excel_range(node)
        if excel_range is None:
            return Unsupported("multi-sheet or malformed range")
        return excel_range.resolve(get_cell_value)

    if isinstance(node, UnaryOpNode):
        value = _eval(node.operand, get_cell_value, functions, max_depth, depth=depth + 1)
        if isinstance(value, Unsupported):
            return value
        if isinstance(value, XlError):
            return value
        if node.op == "-":
            return xl_neg(value)
        if node.op == "%":
            return xl_percent(value)
        return Unsupported(f"Unsupported unary operator {node.op!r}")

    if isinstance(node, BinaryOpNode):
        left = _eval(node.left, get_cell_value, functions, max_depth, depth=depth + 1)
        if isinstance(left, Unsupported):
            return left
        if isinstance(left, XlError):
            return left

        right = _eval(node.right, get_cell_value, functions, max_depth, depth=depth + 1)
        if isinstance(right, Unsupported):
            return right
        if isinstance(right, XlError):
            return right

        op = node.op
        if op == "+":
            return xl_add(left, right)
        if op == "-":
            return xl_sub(left, right)
        if op == "*":
            return xl_mul(left, right)
        if op == "/":
            return xl_div(left, right)
        if op == "^":
            return xl_pow(left, right)
        if op == "&":
            return xl_concat(left, right)
        if op == "=":
            return xl_eq(left, right)
        if op == "<":
            return xl_lt(left, right)
        if op == ">":
            return xl_gt(left, right)
        if op == "<=":
            return xl_le(left, right)
        if op == ">=":
            return xl_ge(left, right)
        if op == "<>":
            return xl_ne(left, right)
        return Unsupported(f"Unsupported binary operator {op!r}")

    if isinstance(node, FunctionCallNode):
        args: list[CellValue | XlError | Unsupported] = [
            _eval(arg, get_cell_value, functions, max_depth, depth=depth + 1)
            for arg in node.args
        ]
        for value in args:
            if isinstance(value, Unsupported):
                return value
            if isinstance(value, XlError):
                return value

        flat_args: list[CellValue] = [
            cast(CellValue, v)
            for v in args
            if not isinstance(v, (XlError, Unsupported))
        ]

        name = node.name.upper()
        impl = functions.get(name) or _DEFAULT_FUNCTIONS.get(name)
        if impl is None:
            return Unsupported(f"Unsupported function {name!r}")
        return impl(flat_args)

    return Unsupported(f"Unsupported AST node type {type(node).__name__}")


def _range_node_to_excel_range(node: RangeNode) -> ExcelRange | None:
    """Convert a RangeNode into an ExcelRange, if it refers to a single sheet."""

    try:
        sheet_start, coord_start = node.start.split("!", 1)
        sheet_end, coord_end = node.end.split("!", 1)
    except ValueError:
        return None

    if sheet_start != sheet_end:
        return None

    row1, col1 = coordinate_to_tuple(coord_start)
    row2, col2 = coordinate_to_tuple(coord_end)

    start_row, end_row = sorted((row1, row2))
    start_col, end_col = sorted((col1, col2))

    return ExcelRange(
        sheet=sheet_start,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )


def _fn_sum(args: list[CellValue]) -> CellValue:
    nums, err = numeric_values(flatten(*args))
    if err is not None:
        return err
    return float(sum(nums))


def _fn_min(args: list[CellValue]) -> CellValue:
    nums, err = numeric_values(flatten(*args))
    if err is not None:
        return err
    if not nums:
        return XlError.VALUE
    return float(min(nums))


def _fn_max(args: list[CellValue]) -> CellValue:
    nums, err = numeric_values(flatten(*args))
    if err is not None:
        return err
    if not nums:
        return XlError.VALUE
    return float(max(nums))


def _fn_abs(args: list[CellValue]) -> CellValue:
    if not args:
        return XlError.VALUE
    n = to_number(args[0])
    if isinstance(n, XlError):
        return n
    return float(abs(n))


def _fn_if(args: list[CellValue]) -> CellValue:
    if len(args) < 2:
        return XlError.VALUE

    cond = to_bool(args[0])
    if isinstance(cond, XlError):
        return cond

    if cond:
        return args[1]

    if len(args) >= 3:
        return args[2]

    # Excel treats missing else as FALSE in a boolean context; here we surface False.
    return False


_DEFAULT_FUNCTIONS: dict[str, Callable[[list[CellValue]], CellValue]] = {
    "SUM": _fn_sum,
    "MIN": _fn_min,
    "MAX": _fn_max,
    "ABS": _fn_abs,
    "IF": _fn_if,
}

