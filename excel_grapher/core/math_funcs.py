"""Shared math worksheet function implementations."""

from __future__ import annotations

import math
import re
from typing import TypeVar

import numpy as np

from .coercions import (
    excel_casefold,
    flatten,
    numeric_values,
    to_bool,
    to_number,
    to_string,
)
from .types import CellValue, XlError

T = TypeVar("T", str, float)

__all__ = [
    "abs_number",
    "average_cells",
    "averageif_cells",
    "countif_cells",
    "exp_number",
    "large_kth",
    "max_cells",
    "min_cells",
    "normdist_value",
    "npv_cells",
    "rank_number",
    "round_number",
    "rounddown_number",
    "stdev_cells",
    "sum_cells",
]


def sum_cells(*args: CellValue) -> float | XlError:
    """Return the sum of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    return float(sum(nums))


def average_cells(*args: CellValue) -> float | XlError:
    """Return the average of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return XlError.DIV
    return float(sum(nums) / len(nums))


def min_cells(*args: CellValue) -> float | XlError:
    """Return the minimum of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(min(nums))


def max_cells(*args: CellValue) -> float | XlError:
    """Return the maximum of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(max(nums))


def round_number(number: CellValue, num_digits: CellValue) -> float | XlError:
    """Round a number to the given number of digits."""
    n = to_number(number)
    if isinstance(n, XlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, XlError):
        return d
    return float(round(n, int(d)))


def rounddown_number(number: CellValue, num_digits: CellValue) -> float | XlError:
    """Round a number down to the given number of digits."""
    n = to_number(number)
    if isinstance(n, XlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, XlError):
        return d
    digits = int(d)
    factor = 10**digits
    if n >= 0:
        return float(math.floor(n * factor) / factor)
    return float(math.ceil(n * factor) / factor)


def npv_cells(rate: CellValue, *values: CellValue) -> float | XlError:
    """Return net present value for a rate and cash flows."""
    r = to_number(rate)
    if isinstance(r, XlError):
        return r
    all_values = list(flatten(*values))
    nums, err = numeric_values(all_values)
    if err is not None:
        return err
    if len(nums) == 0:
        return XlError.VALUE
    result = 0.0
    for i, val in enumerate(nums):
        result += val / ((1 + r) ** (i + 1))
    return result


def stdev_cells(*args: CellValue) -> float | XlError:
    """Return sample standard deviation of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) < 2:
        return XlError.DIV
    mean = sum(nums) / len(nums)
    variance = sum((x - mean) ** 2 for x in nums) / (len(nums) - 1)
    return float(variance**0.5)


def _iter_numeric_cells(values: list[CellValue]) -> tuple[list[float], XlError | None]:
    nums: list[float] = []
    for v in values:
        if isinstance(v, XlError):
            return ([], v)
        if v is None:
            continue
        if isinstance(v, bool):
            continue
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            nums.append(float(v))
            continue
        if isinstance(v, (np.integer, np.floating)):
            nums.append(float(v))
            continue
    return (nums, None)


def _wildcard_to_regex(pattern: str) -> re.Pattern[str]:
    out: list[str] = ["^"]
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern):
            i += 1
            out.append(re.escape(pattern[i]))
        elif ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    out.append("$")
    return re.compile("".join(out), re.IGNORECASE)


def _parse_countif_criteria(criteria: str) -> tuple[str | None, str]:
    s = criteria.strip()
    for op in (">=", "<=", "<>", ">", "<", "="):
        if s.startswith(op):
            return (op, s[len(op) :].strip())
    return (None, s)


def _criteria_compare(op: str, left: T, right: T) -> bool:
    """Compare two values of the same type."""
    if op == "=":
        return left == right
    if op == "<>":
        return left != right
    if op == ">":
        return left > right
    if op == "<":
        return left < right
    if op == ">=":
        return left >= right
    if op == "<=":
        return left <= right
    return False


def _value_matches_criteria(cell_value: CellValue, criteria: CellValue) -> bool:
    if isinstance(criteria, XlError):
        return False
    if not isinstance(criteria, str):
        target = criteria
        if isinstance(cell_value, XlError):
            return False
        if target is None:
            return cell_value is None
        if isinstance(target, bool):
            b = to_bool(cell_value)
            return (not isinstance(b, XlError)) and b == target
        if isinstance(target, (int, float)) and not isinstance(target, bool):
            vn = to_number(cell_value)
            return (not isinstance(vn, XlError)) and vn == float(target)
        return excel_casefold(to_string(cell_value)) == excel_casefold(to_string(target))

    op, rhs = _parse_countif_criteria(criteria)
    if isinstance(cell_value, XlError):
        return False

    if op is None:
        if any(ch in rhs for ch in ("*", "?", "~")):
            rx = _wildcard_to_regex(rhs)
            return rx.match(to_string(cell_value)) is not None
        return excel_casefold(to_string(cell_value)) == excel_casefold(rhs)

    try:
        rhs_num = float(rhs) if rhs != "" else 0.0
    except ValueError:
        rhs_num = None

    if rhs_num is not None:
        vn = to_number(cell_value)
        if isinstance(vn, XlError):
            return False
        return _criteria_compare(op, vn, rhs_num)

    return _criteria_compare(op, excel_casefold(to_string(cell_value)), excel_casefold(rhs))


def countif_cells(range_values: CellValue, criteria: CellValue) -> int | XlError:
    """Count cells matching criteria."""
    if isinstance(criteria, XlError):
        return criteria
    values = list(flatten(range_values))
    return sum(1 for v in values if _value_matches_criteria(v, criteria))


def averageif_cells(
    range_values: CellValue,
    criteria: CellValue,
    average_range: CellValue | None = None,
) -> float | XlError:
    """Return the average of cells matching criteria."""
    if isinstance(criteria, XlError):
        return criteria
    crit_vals = list(flatten(range_values))
    avg_vals = list(flatten(average_range if average_range is not None else range_values))
    if len(crit_vals) != len(avg_vals):
        return XlError.VALUE
    matched: list[float] = []
    for crit_val, avg_val in zip(crit_vals, avg_vals, strict=True):
        if not _value_matches_criteria(crit_val, criteria):
            continue
        n = to_number(avg_val)
        if isinstance(n, XlError):
            return n
        matched.append(float(n))
    if not matched:
        return XlError.DIV
    return sum(matched) / len(matched)


def large_kth(array: CellValue, k: CellValue) -> float | XlError:
    """Return the k-th largest numeric value."""
    kk = to_number(k)
    if isinstance(kk, XlError):
        return kk
    kth = int(kk)
    if kth < 1:
        return XlError.NUM
    values = list(flatten(array))
    nums, err = _iter_numeric_cells(values)
    if err is not None:
        return err
    if kth > len(nums):
        return XlError.NUM
    nums.sort(reverse=True)
    return float(nums[kth - 1])


def rank_number(number: CellValue, ref: CellValue, order: CellValue = 0) -> int | XlError:
    """Return the rank of a number within a reference range."""
    nn = to_number(number)
    if isinstance(nn, XlError):
        return nn
    oo = to_number(order)
    if isinstance(oo, XlError):
        return oo
    ascending = int(oo) != 0
    values = list(flatten(ref))
    nums, err = _iter_numeric_cells(values)
    if err is not None:
        return err
    if ascending:
        return 1 + sum(1 for v in nums if v < nn)
    return 1 + sum(1 for v in nums if v > nn)


def normdist_value(
    x: CellValue,
    mean: CellValue,
    standard_dev: CellValue,
    cumulative: CellValue,
) -> float | XlError:
    """Return the normal distribution value."""
    xx = to_number(x)
    if isinstance(xx, XlError):
        return xx
    mm = to_number(mean)
    if isinstance(mm, XlError):
        return mm
    sd = to_number(standard_dev)
    if isinstance(sd, XlError):
        return sd
    if sd <= 0:
        return XlError.NUM
    cc = to_bool(cumulative)
    if isinstance(cc, XlError):
        return cc
    z = (xx - mm) / sd
    if cc:
        return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))
    return (1.0 / (sd * math.sqrt(2.0 * math.pi))) * math.exp(-0.5 * z * z)


def abs_number(*args: CellValue) -> float | XlError:
    """Return the absolute value of a number (Excel ``ABS``)."""
    if len(args) != 1:
        return XlError.VALUE
    n = to_number(args[0])
    if isinstance(n, XlError):
        return n
    return float(abs(n))


def exp_number(*args: CellValue) -> float | XlError:
    """Return e raised to the power of a number (Excel ``EXP``)."""
    if len(args) != 1:
        return XlError.VALUE
    n = to_number(args[0])
    if isinstance(n, XlError):
        return n
    try:
        return float(math.exp(n))
    except OverflowError:
        return XlError.NUM
