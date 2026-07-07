"""Raise-only boundary wrappers for math worksheet functions."""

from __future__ import annotations

from excel_grapher.core import CellValue
from excel_grapher.core.math_funcs import (
    abs_number,
    average_cells,
    averageif_cells,
    countif_cells,
    exp_number,
    large_kth,
    max_cells,
    min_cells,
    normdist_value,
    npv_cells,
    rank_number,
    round_number,
    rounddown_number,
    stdev_cells,
    sum_cells,
)

from .errors import raise_if_sentinel_float, raise_if_sentinel_int

__all__ = [
    "xl_abs",
    "xl_average",
    "xl_averageif",
    "xl_countif",
    "xl_exp",
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
]


def xl_sum(*args: CellValue) -> float:
    """Return the sum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(sum_cells(*args))


def xl_average(*args: CellValue) -> float:
    """Return the average of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(average_cells(*args))


def xl_min(*args: CellValue) -> float:
    """Return the minimum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(min_cells(*args))


def xl_max(*args: CellValue) -> float:
    """Return the maximum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(max_cells(*args))


def xl_round(number: CellValue, num_digits: CellValue) -> float:
    """Round a number, raising on Excel coercion errors."""
    return raise_if_sentinel_float(round_number(number, num_digits))


def xl_rounddown(number: CellValue, num_digits: CellValue) -> float:
    """Round a number down, raising on Excel coercion errors."""
    return raise_if_sentinel_float(rounddown_number(number, num_digits))


def xl_npv(rate: CellValue, *values: CellValue) -> float:
    """Return net present value, raising on Excel errors."""
    return raise_if_sentinel_float(npv_cells(rate, *values))


def xl_stdev(*args: CellValue) -> float:
    """Return sample standard deviation, raising on Excel errors."""
    return raise_if_sentinel_float(stdev_cells(*args))


def xl_countif(range_values: CellValue, criteria: CellValue) -> int:
    """Count cells matching criteria, raising on Excel errors."""
    return raise_if_sentinel_int(countif_cells(range_values, criteria))


def xl_averageif(
    range_values: CellValue,
    criteria: CellValue,
    average_range: CellValue | None = None,
) -> float:
    """Return the average of cells matching criteria, raising on Excel errors."""
    return raise_if_sentinel_float(averageif_cells(range_values, criteria, average_range))


def xl_large(array: CellValue, k: CellValue) -> float:
    """Return the k-th largest value, raising on Excel errors."""
    return raise_if_sentinel_float(large_kth(array, k))


def xl_rank(number: CellValue, ref: CellValue, order: CellValue = 0) -> int:
    """Return the rank of a number in a list, raising on Excel errors."""
    return raise_if_sentinel_int(rank_number(number, ref, order))


def xl_normdist(
    x: CellValue,
    mean: CellValue,
    standard_dev: CellValue,
    cumulative: CellValue,
) -> float:
    """Return the normal distribution value, raising on Excel errors."""
    return raise_if_sentinel_float(normdist_value(x, mean, standard_dev, cumulative))


def xl_abs(*args: CellValue) -> float:
    """Return the absolute value of a number, raising on Excel errors."""
    return raise_if_sentinel_float(abs_number(*args))


def xl_exp(*args: CellValue) -> float:
    """Return e raised to a power, raising on Excel errors."""
    return raise_if_sentinel_float(exp_number(*args))
