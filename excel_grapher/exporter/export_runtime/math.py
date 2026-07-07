"""Raise-only boundary wrappers for math runtime helpers."""

from __future__ import annotations

from excel_grapher.core import CellValue

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
    return raise_if_sentinel_float(_sentinel_xl_sum(*args))


def xl_average(*args: CellValue) -> float:
    """Return the average of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_average(*args))


def xl_min(*args: CellValue) -> float:
    """Return the minimum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_min(*args))


def xl_max(*args: CellValue) -> float:
    """Return the maximum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_max(*args))


def xl_round(number: CellValue, num_digits: CellValue) -> float:
    """Round a number, raising on Excel coercion errors."""
    return raise_if_sentinel_float(_sentinel_xl_round(number, num_digits))


def xl_rounddown(number: CellValue, num_digits: CellValue) -> float:
    """Round a number down, raising on Excel coercion errors."""
    return raise_if_sentinel_float(_sentinel_xl_rounddown(number, num_digits))


def xl_npv(rate: CellValue, *values: CellValue) -> float:
    """Return net present value, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_npv(rate, *values))


def xl_stdev(*args: CellValue) -> float:
    """Return sample standard deviation, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_stdev(*args))


def xl_countif(range_values: CellValue, criteria: CellValue) -> int:
    """Count cells matching criteria, raising on Excel errors."""
    return raise_if_sentinel_int(_sentinel_xl_countif(range_values, criteria))


def xl_averageif(
    range_values: CellValue,
    criteria: CellValue,
    average_range: CellValue | None = None,
) -> float:
    """Return the average of cells matching criteria, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_averageif(range_values, criteria, average_range))


def xl_large(array: CellValue, k: CellValue) -> float:
    """Return the k-th largest value, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_large(array, k))


def xl_rank(number: CellValue, ref: CellValue, order: CellValue = 0) -> int:
    """Return the rank of a number in a list, raising on Excel errors."""
    return raise_if_sentinel_int(_sentinel_xl_rank(number, ref, order))


def xl_normdist(
    x: CellValue,
    mean: CellValue,
    standard_dev: CellValue,
    cumulative: CellValue,
) -> float:
    """Return the normal distribution value, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_normdist(x, mean, standard_dev, cumulative))


def xl_abs(*args: CellValue) -> float:
    """Return the absolute value of a number, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_abs(*args))


def xl_exp(*args: CellValue) -> float:
    """Return e raised to a power, raising on Excel errors."""
    return raise_if_sentinel_float(_sentinel_xl_exp(*args))
