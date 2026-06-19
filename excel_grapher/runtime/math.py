from __future__ import annotations

import math
import re
from typing import TypeVar

import numpy as np

from excel_grapher.core import (
    CellValue,
    XlError,
    excel_casefold,
    flatten,
    numeric_values,
    to_bool,
    to_number,
    to_string,
)
from excel_grapher.core.functions import xl_abs

T = TypeVar("T", str, float)

__all__ = [
    "xl_abs",
    "xl_average",
    "xl_averageif",
    "xl_count",
    "xl_counta",
    "xl_countif",
    "xl_large",
    "xl_max",
    "xl_min",
    "xl_normdist",
    "xl_npv",
    "xl_rank",
    "xl_round",
    "xl_rounddown",
    "xl_stdev",
    "xl_sum",
    "xl_sumproduct",
]


def xl_sum(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    return float(sum(nums))


def xl_average(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return XlError.DIV
    return float(sum(nums) / len(nums))


def xl_min(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(min(nums))


def xl_max(*args: CellValue) -> float | XlError:
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(max(nums))


def xl_count(*args: CellValue) -> int:
    count = 0
    for v in flatten(*args):
        if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
            count += 1
    return count


def xl_counta(*args: CellValue) -> int:
    count = 0
    for v in flatten(*args):
        if v is not None and v != "":
            count += 1
    return count


def xl_sumproduct(*args: CellValue) -> float | XlError:
    if len(args) == 0:
        return 0.0
    arrays: list[np.ndarray] = []
    for arg in args:
        if isinstance(arg, np.ndarray):
            arrays.append(arg)
        else:
            arrays.append(np.array([[arg]], dtype=object))
    shape = arrays[0].shape
    for arr in arrays[1:]:
        if arr.shape != shape:
            return XlError.VALUE
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


def xl_round(number: CellValue, num_digits: CellValue) -> float | XlError:
    n = to_number(number)
    if isinstance(n, XlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, XlError):
        return d
    return float(round(n, int(d)))


def xl_rounddown(number: CellValue, num_digits: CellValue) -> float | XlError:
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


def xl_npv(rate: CellValue, *values: CellValue) -> float | XlError:
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


def xl_stdev(*args: CellValue) -> float | XlError:
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


def xl_countif(range_values: CellValue, criteria: CellValue) -> int | XlError:
    if isinstance(criteria, XlError):
        return criteria
    values = list(flatten(range_values))
    return sum(1 for v in values if _value_matches_criteria(v, criteria))


def xl_averageif(
    range_values: CellValue,
    criteria: CellValue,
    average_range: CellValue | None = None,
) -> float | XlError:
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


def _criteria_compare(op: str, left: T, right: T) -> bool:
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


def xl_large(array: CellValue, k: CellValue) -> float | XlError:
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


def xl_rank(number: CellValue, ref: CellValue, order: CellValue = 0) -> int | XlError:
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


def xl_normdist(
    x: CellValue,
    mean: CellValue,
    standard_dev: CellValue,
    cumulative: CellValue,
) -> float | XlError:
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
