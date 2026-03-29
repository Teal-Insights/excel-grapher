"""Excel-style scalar operators (representation-agnostic)."""

from __future__ import annotations

from .coercions import excel_casefold, to_number, to_string
from .types import CellValue, XlError


def _xl_compare(op: str, left: CellValue, right: CellValue) -> bool | XlError:
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


def xl_concat(left: CellValue, right: CellValue) -> str | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    return to_string(left) + to_string(right)


def xl_eq(left: CellValue, right: CellValue) -> bool | XlError:
    return _xl_compare("=", left, right)


def xl_ne(left: CellValue, right: CellValue) -> bool | XlError:
    return _xl_compare("<>", left, right)


def xl_lt(left: CellValue, right: CellValue) -> bool | XlError:
    return _xl_compare("<", left, right)


def xl_gt(left: CellValue, right: CellValue) -> bool | XlError:
    return _xl_compare(">", left, right)


def xl_le(left: CellValue, right: CellValue) -> bool | XlError:
    return _xl_compare("<=", left, right)


def xl_ge(left: CellValue, right: CellValue) -> bool | XlError:
    return _xl_compare(">=", left, right)


def xl_iferror(value: CellValue, value_if_error: CellValue) -> CellValue:
    if isinstance(value, XlError):
        return value_if_error
    return value


def xl_div(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    if rn == 0:
        return XlError.DIV
    return ln / rn


def xl_add(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return ln + rn


def xl_sub(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return ln - rn


def xl_mul(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    return ln * rn


def xl_pow(left: CellValue, right: CellValue) -> float | XlError:
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right
    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError):
        return ln
    if isinstance(rn, XlError):
        return rn
    try:
        return ln**rn
    except (ValueError, OverflowError):
        return XlError.NUM


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
