"""Excel operators over scalars and lazy ranges for exported code."""

from __future__ import annotations

from excel_grapher.core import XlError, to_number
from excel_grapher.core.operators_reference import (
    apply_arithmetic,
    compare_scalars,
    concat_scalars,
)

from .values import CellValue, Grid, as_scalar

__all__ = [
    "xl_add",
    "xl_concat",
    "xl_div",
    "xl_eq",
    "xl_ge",
    "xl_gt",
    "xl_le",
    "xl_lt",
    "xl_mul",
    "xl_ne",
    "xl_neg",
    "xl_percent",
    "xl_pos",
    "xl_pow",
    "xl_sub",
]


def _broadcast_pair(left: CellValue, right: CellValue) -> tuple[Grid, Grid] | XlError | None:
    """Wrap array operands as aligned grids; `None` when both operands are scalar."""
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is None and right_grid is None:
        return None
    if left_grid is not None and right_grid is not None:
        if (left_grid.nrows, left_grid.ncols) != (right_grid.nrows, right_grid.ncols):
            return XlError.VALUE
        return left_grid, right_grid
    if left_grid is not None:
        scalar_right = Grid.wrap([[right] * left_grid.ncols for _ in range(left_grid.nrows)])
        assert scalar_right is not None
        return left_grid, scalar_right
    assert right_grid is not None
    scalar_left = Grid.wrap([[left] * right_grid.ncols for _ in range(right_grid.nrows)])
    assert scalar_left is not None
    return scalar_left, right_grid


def _xl_compare(op: str, left: CellValue, right: CellValue) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    pair = _broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    if pair is None:
        return compare_scalars(op, as_scalar(left), as_scalar(right))

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            cell = compare_scalars(op, arr_left.at(row0, col0), arr_right.at(row0, col0))
            if isinstance(cell, XlError):
                return cell
            out_row.append(cell)
        result.append(out_row)
    return result


def _xl_arithmetic(op: str, left: CellValue, right: CellValue) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    pair = _broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    if pair is None:
        ln = to_number(as_scalar(left))
        rn = to_number(as_scalar(right))
        if isinstance(ln, XlError):
            return ln
        if isinstance(rn, XlError):
            return rn
        return apply_arithmetic(op, ln, rn)

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            ln = to_number(arr_left.at(row0, col0))
            if isinstance(ln, XlError):
                return ln
            rn = to_number(arr_right.at(row0, col0))
            if isinstance(rn, XlError):
                return rn
            cell = apply_arithmetic(op, ln, rn)
            if isinstance(cell, XlError):
                return cell
            out_row.append(cell)
        result.append(out_row)
    return result


def xl_concat(left: CellValue, right: CellValue) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    pair = _broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    if pair is None:
        return concat_scalars(as_scalar(left), as_scalar(right))

    arr_left, arr_right = pair
    return [
        [
            concat_scalars(arr_left.at(row0, col0), arr_right.at(row0, col0))
            for col0 in range(arr_left.ncols)
        ]
        for row0 in range(arr_left.nrows)
    ]


def xl_eq(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("=", left, right)


def xl_ne(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("<>", left, right)


def xl_lt(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("<", left, right)


def xl_gt(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare(">", left, right)


def xl_le(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare("<=", left, right)


def xl_ge(left: CellValue, right: CellValue) -> CellValue:
    return _xl_compare(">=", left, right)


def xl_div(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("/", left, right)


def xl_add(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("+", left, right)


def xl_sub(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("-", left, right)


def xl_mul(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("*", left, right)


def xl_pow(left: CellValue, right: CellValue) -> CellValue:
    return _xl_arithmetic("^", left, right)


def xl_neg(value: CellValue) -> float | XlError:
    if isinstance(value, XlError):
        return value
    n = to_number(as_scalar(value))
    if isinstance(n, XlError):
        return n
    return -n


def xl_pos(value: CellValue) -> float | XlError:
    if isinstance(value, XlError):
        return value
    n = to_number(as_scalar(value))
    if isinstance(n, XlError):
        return n
    return +n


def xl_percent(value: CellValue) -> float | XlError:
    """Excel postfix percent operator (%): divide a numeric value by 100."""
    if isinstance(value, XlError):
        return value
    n = to_number(as_scalar(value))
    if isinstance(n, XlError):
        return n
    return n / 100.0
