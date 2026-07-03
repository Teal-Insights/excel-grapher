"""Excel operators for exported code: coercion helpers and array broadcast maps.

Scalar formulas emit native Python operators in generated code; these helpers
cover coercion, Excel-specific error semantics, and lazy-range array broadcast.
"""

from __future__ import annotations

from excel_grapher.core import XlError, to_bool, to_int, to_number
from excel_grapher.core.operators_reference import (
    apply_arithmetic,
    compare_scalars,
    concat_scalars,
)
from excel_grapher.core.types import XlErrorException

from .ranges import Range
from .values import CellValue, ExcelRange, Grid, Scalar, as_scalar

__all__ = [
    "xl_compare",
    "xl_bool",
    "xl_int",
    "xl_is_array",
    "xl_map_arithmetic",
    "xl_map_compare",
    "xl_map_concat",
    "xl_map_unary",
    "xl_number",
    "xl_pow_numbers",
]


def _raise_error(code: XlError) -> XlErrorException:
    """Build the exception for an Excel error code (callers raise the result)."""
    return XlErrorException(code)


def xl_number(value: CellValue) -> float:
    """Coerce a scalar cell value to a number, raising on Excel errors."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    number = to_number(scalar)
    if isinstance(number, XlError):
        raise _raise_error(number)
    return number


def xl_int(value: CellValue) -> int:
    """Coerce a scalar cell value to an integer, raising on Excel errors."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    integer = to_int(scalar)
    if isinstance(integer, XlError):
        raise _raise_error(integer)
    return integer


def xl_bool(value: CellValue) -> bool:
    """Coerce a scalar cell value to a boolean, raising on Excel errors."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    boolean = to_bool(scalar)
    if isinstance(boolean, XlError):
        raise _raise_error(boolean)
    return boolean


def xl_is_array(value: object) -> bool:
    """Return whether *value* is a range or nested-list array operand."""
    return isinstance(value, (Range, ExcelRange, list, tuple))


def xl_pow_numbers(left: float, right: float) -> float:
    """Apply Excel exponentiation to coerced numbers."""
    try:
        value = left**right
    except (ValueError, OverflowError):
        raise _raise_error(XlError.NUM) from None
    if isinstance(value, complex):
        raise _raise_error(XlError.NUM)
    return value


def xl_compare(op: str, left: CellValue, right: CellValue) -> bool:
    """Compare two scalar operands with Excel ordering rules."""
    result = compare_scalars(op, as_scalar(left), as_scalar(right))
    if isinstance(result, XlError):
        raise _raise_error(result)
    return result


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


def _apply_arithmetic_or_raise(op: str, left: CellValue, right: CellValue) -> CellValue:
    ln = xl_number(left)
    rn = xl_number(right)
    if op == "^":
        return xl_pow_numbers(ln, rn)
    cell = apply_arithmetic(op, ln, rn)
    if isinstance(cell, XlError):
        raise _raise_error(cell)
    return cell


def xl_map_arithmetic(op: str, left: CellValue, right: CellValue) -> CellValue:
    """Element-wise arithmetic over scalar or broadcast array operands."""
    pair = _broadcast_pair(left, right)
    if pair is None:
        return _apply_arithmetic_or_raise(op, left, right)

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            out_row.append(
                _apply_arithmetic_or_raise(
                    op, _cell_or_raise(arr_left, row0, col0), _cell_or_raise(arr_right, row0, col0)
                )
            )
        result.append(out_row)
    return result


def xl_map_compare(op: str, left: CellValue, right: CellValue) -> CellValue:
    """Element-wise comparison over scalar or broadcast array operands."""
    pair = _broadcast_pair(left, right)
    if pair is None:
        return xl_compare(op, left, right)

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


def xl_map_concat(left: CellValue, right: CellValue) -> CellValue:
    """Element-wise string concatenation over scalar or broadcast array operands."""
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


def xl_map_unary(op: str, value: CellValue) -> CellValue:
    """Apply a unary Excel operator over scalar or array operands."""
    grid = Grid.wrap(value)
    if grid is None:
        number = xl_number(value)
        if op == "-":
            return -number
        if op == "+":
            return +number
        if op == "%":
            return number / 100.0
        raise ValueError(f"Unknown unary operator: {op}")

    result: list[list[CellValue]] = []
    for row0 in range(grid.nrows):
        out_row: list[CellValue] = []
        for col0 in range(grid.ncols):
            number = xl_number(_cell_or_raise(grid, row0, col0))
            if op == "-":
                out_row.append(-number)
            elif op == "+":
                out_row.append(+number)
            elif op == "%":
                out_row.append(number / 100.0)
            else:
                raise ValueError(f"Unknown unary operator: {op}")
        result.append(out_row)
    return result
