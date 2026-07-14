"""Excel operators for exported code: coercion helpers and array broadcast maps.

Scalar formulas emit native Python operators in generated code; these helpers
cover coercion, Excel-specific error semantics, and lazy-range array broadcast.

Element-wise maps delegate to shared `excel_grapher.core.operator_maps` (sentinel
returns) and raise `XlErrorException` at this boundary.
"""

from __future__ import annotations

from typing import cast

from excel_grapher.core import XlError, to_bool, to_int, to_number
from excel_grapher.core.operator_maps import (
    map_arithmetic,
    map_compare,
    map_concat,
    map_unary,
)
from excel_grapher.core.operators_reference import compare_scalars
from excel_grapher.core.types import XlErrorException

from .ranges import Range
from .values import CellValue, ExcelRange, as_scalar

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


def _raise_if_error(value: object) -> CellValue:
    if isinstance(value, XlError):
        raise _raise_error(value)
    return cast(CellValue, value)


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


def xl_map_arithmetic(op: str, left: CellValue, right: CellValue) -> CellValue:
    """Element-wise arithmetic over scalar or broadcast array operands."""
    return _raise_if_error(map_arithmetic(op, left, right))


def xl_map_compare(op: str, left: CellValue, right: CellValue) -> CellValue:
    """Element-wise comparison over scalar or broadcast array operands."""
    return _raise_if_error(map_compare(op, left, right))


def xl_map_concat(left: CellValue, right: CellValue) -> CellValue:
    """Element-wise string concatenation over scalar or broadcast array operands."""
    return _raise_if_error(map_concat(left, right))


def xl_map_unary(op: str, value: CellValue) -> CellValue:
    """Apply a unary Excel operator over scalar or array operands."""
    return _raise_if_error(map_unary(op, value))
