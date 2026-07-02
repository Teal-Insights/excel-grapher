"""Excel operators over scalars and lazy ranges for exported code."""

from __future__ import annotations

from excel_grapher.core import XlError, to_number
from excel_grapher.core.operators_reference import (
    apply_arithmetic,
    compare_scalars,
    concat_scalars,
)
from excel_grapher.core.types import XlErrorException

from .values import CellValue, Grid, Scalar, as_scalar

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


def _raise_error(code: XlError) -> XlErrorException:
    """Build the exception for an Excel error code (callers raise the result)."""
    return XlErrorException(code)


def _scalar_or_raise(value: CellValue) -> Scalar:
    """Collapse to a scalar, raising when the value is an Excel error."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    return scalar


def _cell_or_raise(grid: Grid, row0: int, col0: int) -> Scalar:
    """Read one grid cell, raising when the stored value is an error sentinel."""
    value = grid.at(row0, col0)
    if isinstance(value, XlError):
        raise _raise_error(value)
    return value


def _broadcast_pair(left: CellValue, right: CellValue) -> tuple[Grid, Grid] | None:
    """Wrap array operands as aligned grids; `None` when both operands are scalar."""
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is None and right_grid is None:
        return None
    if left_grid is not None and right_grid is not None:
        if (left_grid.nrows, left_grid.ncols) != (right_grid.nrows, right_grid.ncols):
            raise _raise_error(XlError.VALUE)
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
    pair = _broadcast_pair(left, right)
    if pair is None:
        cell = compare_scalars(op, _scalar_or_raise(left), _scalar_or_raise(right))
        if isinstance(cell, XlError):
            raise _raise_error(cell)
        return cell

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            cell = compare_scalars(
                op, _cell_or_raise(arr_left, row0, col0), _cell_or_raise(arr_right, row0, col0)
            )
            if isinstance(cell, XlError):
                raise _raise_error(cell)
            out_row.append(cell)
        result.append(out_row)
    return result


def _number_or_raise(value: Scalar) -> float:
    """Coerce a scalar to a number, raising on Excel coercion errors."""
    number = to_number(value)
    if isinstance(number, XlError):
        raise _raise_error(number)
    return number


def _xl_arithmetic(op: str, left: CellValue, right: CellValue) -> CellValue:
    pair = _broadcast_pair(left, right)
    if pair is None:
        ln = _number_or_raise(_scalar_or_raise(left))
        rn = _number_or_raise(_scalar_or_raise(right))
        cell = apply_arithmetic(op, ln, rn)
        if isinstance(cell, XlError):
            raise _raise_error(cell)
        return cell

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            ln = _number_or_raise(_cell_or_raise(arr_left, row0, col0))
            rn = _number_or_raise(_cell_or_raise(arr_right, row0, col0))
            cell = apply_arithmetic(op, ln, rn)
            if isinstance(cell, XlError):
                raise _raise_error(cell)
            out_row.append(cell)
        result.append(out_row)
    return result


def xl_concat(left: CellValue, right: CellValue) -> CellValue:
    pair = _broadcast_pair(left, right)
    if pair is None:
        return concat_scalars(_scalar_or_raise(left), _scalar_or_raise(right))

    arr_left, arr_right = pair
    return [
        [
            concat_scalars(
                _cell_or_raise(arr_left, row0, col0), _cell_or_raise(arr_right, row0, col0)
            )
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


def xl_neg(value: CellValue) -> float:
    return -_number_or_raise(_scalar_or_raise(value))


def xl_pos(value: CellValue) -> float:
    return +_number_or_raise(_scalar_or_raise(value))


def xl_percent(value: CellValue) -> float:
    """Excel postfix percent operator (%): divide a numeric value by 100."""
    return _number_or_raise(_scalar_or_raise(value)) / 100.0
