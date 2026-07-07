"""Raise-only boundary wrappers for text worksheet functions."""

from __future__ import annotations

from excel_grapher.core import CellValue
from excel_grapher.core.text_funcs import (
    concatenate_cells,
    left_chars,
    lower_text,
    mid_chars,
    numbervalue_parse,
    right_chars,
    text_format,
    value_from_text,
)

from .errors import raise_if_sentinel_float, raise_if_sentinel_str

__all__ = [
    "xl_concatenate",
    "xl_left",
    "xl_lower",
    "xl_mid",
    "xl_numbervalue",
    "xl_right",
    "xl_text",
    "xl_value",
]


def xl_left(text: CellValue, num_chars: CellValue = 1) -> str:
    """Return the leftmost characters of text, raising on Excel errors."""
    return raise_if_sentinel_str(left_chars(text, num_chars))


def xl_right(text: CellValue, num_chars: CellValue = 1) -> str:
    """Return the rightmost characters of text, raising on Excel errors."""
    return raise_if_sentinel_str(right_chars(text, num_chars))


def xl_mid(text: CellValue, start_num: CellValue, num_chars: CellValue) -> str:
    """Return characters from the middle of text, raising on Excel errors."""
    return raise_if_sentinel_str(mid_chars(text, start_num, num_chars))


def xl_concatenate(*args: CellValue) -> str:
    """Concatenate text values, raising on Excel errors."""
    return raise_if_sentinel_str(concatenate_cells(*args))


def xl_text(value: CellValue, format_text: CellValue) -> str:
    """Format a value as text, raising on Excel errors."""
    return raise_if_sentinel_str(text_format(value, format_text))


def xl_numbervalue(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float:
    """Convert text to a number, raising on Excel errors."""
    return raise_if_sentinel_float(numbervalue_parse(text, decimal_separator, group_separator))


def xl_lower(text: CellValue) -> str:
    """Return lowercase text, raising on Excel errors."""
    return raise_if_sentinel_str(lower_text(text))


def xl_value(text: CellValue) -> float:
    """Convert text to a number, preserving Excel ``VALUE`` fallback semantics."""
    return raise_if_sentinel_float(value_from_text(text))
