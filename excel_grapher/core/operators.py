"""Excel-style scalar and array operators (representation-agnostic)."""

from __future__ import annotations

import numpy as np

from .coercions import excel_casefold, to_number, to_string
from .types import CellValue, XlError


def _broadcast_pair(
    left: CellValue,
    right: CellValue,
) -> tuple[np.ndarray, np.ndarray] | XlError:
    """Broadcast scalar/array operands to matching object ndarrays."""
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    if isinstance(left, np.ndarray) and isinstance(right, np.ndarray):
        if left.shape != right.shape:
            return XlError.VALUE
        return left, right
    if isinstance(left, np.ndarray):
        return left, np.full(left.shape, right, dtype=object)
    if isinstance(right, np.ndarray):
        return np.full(right.shape, left, dtype=object), right
    raise TypeError("expected at least one ndarray operand")


def _compare_scalars(op: str, left: CellValue, right: CellValue) -> bool | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    def _cmp_str(a: str, b: str) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    def _cmp_float(a: float, b: float) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    if isinstance(left, str) and isinstance(right, str):
        return _cmp_str(excel_casefold(left), excel_casefold(right))

    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError) or isinstance(rn, XlError):
        return _cmp_str(excel_casefold(to_string(left)), excel_casefold(to_string(right)))

    return _cmp_float(float(ln), float(rn))


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
        result = np.empty(arr_left.shape, dtype=object)
        for indices in np.ndindex(arr_left.shape):
            cell = _compare_scalars(op, arr_left[indices], arr_right[indices])
            # Fail-fast: first cell error replaces the whole array (matches xl_sumproduct).
            if isinstance(cell, XlError):
                return cell
            result[indices] = cell
        return result

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
        result = np.empty(arr_left.shape, dtype=object)
        for indices in np.ndindex(arr_left.shape):
            ln = to_number(arr_left[indices])
            rn = to_number(arr_right[indices])
            if isinstance(ln, XlError):
                return ln
            if isinstance(rn, XlError):
                return rn
            # Fail-fast per cell (aligned with xl_sumproduct error propagation).
            if op == "+":
                result[indices] = ln + rn
            elif op == "-":
                result[indices] = ln - rn
            elif op == "*":
                result[indices] = ln * rn
            elif op == "/":
                if rn == 0:
                    return XlError.DIV
                result[indices] = ln / rn
            elif op == "^":
                try:
                    value = ln**rn
                except (ValueError, OverflowError):
                    return XlError.NUM
                if isinstance(value, complex):
                    return XlError.NUM
                result[indices] = value
            else:
                raise ValueError(f"Unknown arithmetic operator: {op}")
        return result

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
    return to_string(left) + to_string(right)


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
        result = np.empty(arr_left.shape, dtype=object)
        for indices in np.ndindex(arr_left.shape):
            result[indices] = _concat_scalars(arr_left[indices], arr_right[indices])
        return result

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
