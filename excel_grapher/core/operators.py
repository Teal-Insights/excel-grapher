"""Excel-style scalar and array operators (representation-agnostic).

Array operands try vectorized fast paths in ``operators_fastpath`` first; when
size or cell-type guards fail, per-cell reference loops in ``operators_reference``
preserve Excel coercion, broadcasting, and error semantics.
"""

from __future__ import annotations

import numpy as np

from .coercions import to_number
from .operators_fastpath import (
    try_fastpath_arithmetic_array,
    try_fastpath_compare_array,
    try_fastpath_concat_array,
)
from .operators_reference import (
    broadcast_pair,
    compare_scalars,
    concat_scalars,
    reference_arithmetic_array,
    reference_compare_array,
    reference_concat_array,
)
from .types import CellValue, XlError


def _broadcast_pair(
    left: CellValue,
    right: CellValue,
) -> tuple[np.ndarray, np.ndarray] | XlError:
    return broadcast_pair(left, right)


def _compare_scalars(op: str, left: CellValue, right: CellValue) -> bool | XlError:
    return compare_scalars(op, left, right)


def _xl_compare(op: str, left: CellValue, right: CellValue) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if isinstance(left, np.ndarray) or isinstance(right, np.ndarray):
        pair = _broadcast_pair(left, right)
        if isinstance(pair, XlError):
            return pair
        arr_left, arr_right = pair
        fast = try_fastpath_compare_array(op, arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_compare_array(op, arr_left, arr_right)

    return _compare_scalars(op, left, right)


def _xl_arithmetic(
    op: str,
    left: CellValue,
    right: CellValue,
) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if isinstance(left, np.ndarray) or isinstance(right, np.ndarray):
        pair = _broadcast_pair(left, right)
        if isinstance(pair, XlError):
            return pair
        arr_left, arr_right = pair
        fast = try_fastpath_arithmetic_array(op, arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_arithmetic_array(op, arr_left, arr_right)

    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    if op == "+":
        return ln + rn
    if op == "-":
        return ln - rn
    if op == "*":
        return ln * rn
    if op == "/":
        if rn == 0:
            return XlError.DIV
        return ln / rn
    if op == "^":
        try:
            value = ln**rn
        except (ValueError, OverflowError):
            return XlError.NUM
        if isinstance(value, complex):
            return XlError.NUM
        return value
    raise ValueError(f"Unknown arithmetic operator: {op}")


def _concat_scalars(left: CellValue, right: CellValue) -> str:
    return concat_scalars(left, right)


def _xl_concat(left: CellValue, right: CellValue) -> CellValue:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    if isinstance(left, np.ndarray) or isinstance(right, np.ndarray):
        pair = _broadcast_pair(left, right)
        if isinstance(pair, XlError):
            return pair
        arr_left, arr_right = pair
        fast = try_fastpath_concat_array(arr_left, arr_right)
        if fast is not None:
            return fast
        return reference_concat_array(arr_left, arr_right)

    return _concat_scalars(left, right)


def xl_concat(left: CellValue, right: CellValue) -> CellValue:
    return _xl_concat(left, right)


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


def xl_iferror(value: CellValue, value_if_error: CellValue) -> CellValue:
    if isinstance(value, XlError):
        return value_if_error
    return value


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
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return -n


def xl_pos(value: CellValue) -> float | XlError:
    if isinstance(value, XlError):
        return value
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return +n


def xl_percent(value: CellValue) -> float | XlError:
    """Excel postfix percent operator (%): divide a numeric value by 100."""
    if isinstance(value, XlError):
        return value
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return n / 100.0
