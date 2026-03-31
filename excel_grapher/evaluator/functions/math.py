from __future__ import annotations

# ruff: noqa: I001

import math
import re
from typing import TypeVar

import numpy as np

from . import register
from ..helpers import (
    excel_casefold,
    flatten,
    numeric_values,
    to_bool,
    to_number,
    to_string,
)
from ..types import CellValue, XlError


T = TypeVar("T", str, float)


@register("SUM")
def xl_sum(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    return float(sum(nums))


@register("AVERAGE")
def xl_average(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return XlError.DIV
    return float(sum(nums) / len(nums))


@register("MIN")
def xl_min(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(min(nums))


@register("MAX")
def xl_max(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(max(nums))


@register("COUNT")
def xl_count(*args: CellValue) -> int:
    """Count numeric values only."""
    count = 0
    for v in flatten(*args):
        if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
            count += 1
    return count


@register("COUNTA")
def xl_counta(*args: CellValue) -> int:
    """Count non-empty values."""
    count = 0
    for v in flatten(*args):
        if v is not None and v != "":
            count += 1
    return count


@register("SUMPRODUCT")
def xl_sumproduct(*args: CellValue) -> float | XlError:
    """Multiply corresponding elements and sum the products."""
    if len(args) == 0:
        return 0.0

    # Convert all args to numpy arrays
    arrays: list[np.ndarray] = []
    for arg in args:
        if isinstance(arg, np.ndarray):
            arrays.append(arg)
        else:
            # Single value - wrap in array
            arrays.append(np.array([[arg]], dtype=object))

    # Check all arrays have the same shape
    shape = arrays[0].shape
    for arr in arrays[1:]:
        if arr.shape != shape:
            return XlError.VALUE

    # Convert all values to numbers, multiply, and sum
    result = 0.0
    for indices in np.ndindex(shape):
        product = 1.0
        for arr in arrays:
            val = arr[indices]
            n = to_number(val)
            if isinstance(n, XlError):
                return n
            product *= n
        result += product

    return result


@register("ROUND")
def xl_round(number: CellValue, num_digits: CellValue) -> float | XlError:
    """Round a number to a specified number of digits."""
    n = to_number(number)
    if isinstance(n, XlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, XlError):
        return d
    digits = int(d)
    return float(round(n, digits))


@register("ROUNDDOWN")
def xl_rounddown(number: CellValue, num_digits: CellValue) -> float | XlError:
    """Round a number down (towards zero)."""
    import math

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


@register("NPV")
def xl_npv(rate: CellValue, *values: CellValue) -> float | XlError:
    """Calculate net present value of cash flows at a given discount rate.

    NPV = sum(values[i] / (1 + rate)^(i+1)) for i = 0 to n-1

    Note: Excel's NPV assumes cash flows occur at the END of each period,
    starting from period 1 (not period 0).
    """
    r = to_number(rate)
    if isinstance(r, XlError):
        return r

    # Flatten and get all numeric values
    all_values = list(flatten(*values))
    nums, err = numeric_values(all_values)
    if err is not None:
        return err

    if len(nums) == 0:
        return XlError.VALUE

    # Calculate NPV: sum of value[i] / (1+rate)^(i+1)
    result = 0.0
    for i, val in enumerate(nums):
        result += val / ((1 + r) ** (i + 1))

    return result


@register("STDEV")
def xl_stdev(*args: CellValue) -> float | XlError:
    """Calculate sample standard deviation (ignoring text and logical values)."""
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
    """Yield numeric cells for functions that ignore text and logical values."""
    nums: list[float] = []
    for v in values:
        if isinstance(v, XlError):
            return ([], v)
        if v is None:
            continue
        if isinstance(v, bool):
            # Excel ignores logical values in ranges for these statistical functions.
            continue
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            nums.append(float(v))
            continue
        if isinstance(v, (np.integer, np.floating)):
            nums.append(float(v))
            continue
        # Text and other types are ignored.
    return (nums, None)


def _wildcard_to_regex(pattern: str) -> re.Pattern[str]:
    """Convert Excel COUNTIF wildcard pattern to a case-insensitive regex.

    Supported:
    - '*' matches any sequence
    - '?' matches any single character
    - '~' escapes '*' and '?' and '~'
    """
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


@register("COUNTIF")
def xl_countif(range_values: CellValue, criteria: CellValue) -> int | XlError:
    """Count cells in a range that meet a criterion."""
    if isinstance(criteria, XlError):
        return criteria

    values = list(flatten(range_values))

    # Non-string criteria: compare by string form for broad Excel-like behavior.
    if not isinstance(criteria, str):
        target = criteria

        def pred(v: CellValue) -> bool:
            if isinstance(v, XlError):
                return False
            if target is None:
                return v is None
            if isinstance(target, bool):
                b = to_bool(v)
                return (not isinstance(b, XlError)) and b == target
            if isinstance(target, (int, float)) and not isinstance(target, bool):
                vn = to_number(v)
                return (not isinstance(vn, XlError)) and vn == float(target)
            return excel_casefold(to_string(v)) == excel_casefold(to_string(target))

        return sum(1 for v in values if pred(v))

    op, rhs = _parse_countif_criteria(criteria)

    # Wildcard / equality mode.
    if op is None:
        if any(ch in rhs for ch in ("*", "?", "~")):
            rx = _wildcard_to_regex(rhs)
            return sum(
                1
                for v in values
                if not isinstance(v, XlError) and rx.match(to_string(v)) is not None
            )

        rhs_cf = excel_casefold(rhs)
        return sum(
            1
            for v in values
            if not isinstance(v, XlError) and excel_casefold(to_string(v)) == rhs_cf
        )

    # Operator mode: try numeric compare first if RHS parses as a number.
    rhs_num: float | None
    try:
        rhs_num = float(rhs) if rhs != "" else 0.0
    except ValueError:
        rhs_num = None

    count = 0
    for v in values:
        if isinstance(v, XlError):
            continue

        if rhs_num is not None:
            vn = to_number(v)
            if isinstance(vn, XlError):
                # Non-numeric cells simply don't match numeric criteria.
                continue
            match = _compare_values(op, vn, rhs_num)
        else:
            left_str = excel_casefold(to_string(v))
            right_str = excel_casefold(rhs)
            match = _compare_values(op, left_str, right_str)

        if match:
            count += 1

    return count


def _compare_values(op: str, left: T, right: T) -> bool:
    """Compare two values of the same type."""
    if op == "=":
        return left == right
    if op == "<>":
        return left != right
    if op == ">":
        return left > right  # type: ignore[operator]
    if op == "<":
        return left < right  # type: ignore[operator]
    if op == ">=":
        return left >= right  # type: ignore[operator]
    if op == "<=":
        return left <= right  # type: ignore[operator]
    return False


@register("LARGE")
def xl_large(array: CellValue, k: CellValue) -> float | XlError:
    """Return the k-th largest value in a set of values."""
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


@register("RANK")
def xl_rank(number: CellValue, ref: CellValue, order: CellValue = 0) -> int | XlError:
    """Return the rank of a number in a list of numbers."""
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


@register("NORMDIST")
def xl_normdist(
    x: CellValue,
    mean: CellValue,
    standard_dev: CellValue,
    cumulative: CellValue,
) -> float | XlError:
    """Return the normal distribution for the specified mean and standard deviation."""
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
        # CDF: 0.5 * (1 + erf(z / sqrt(2)))
        return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))

    # PDF: (1 / (sd * sqrt(2*pi))) * exp(-0.5*z^2)
    return (1.0 / (sd * math.sqrt(2.0 * math.pi))) * math.exp(-0.5 * z * z)
