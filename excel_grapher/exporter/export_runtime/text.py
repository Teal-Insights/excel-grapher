"""Raise-only boundary wrappers for text runtime helpers."""

from __future__ import annotations

from excel_grapher.core import CellValue, XlError, to_number

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
    return raise_if_sentinel_str(_sentinel_xl_left(text, num_chars))


def xl_right(text: CellValue, num_chars: CellValue = 1) -> str:
    """Return the rightmost characters of text, raising on Excel errors."""
    return raise_if_sentinel_str(_sentinel_xl_right(text, num_chars))


def xl_mid(text: CellValue, start_num: CellValue, num_chars: CellValue) -> str:
    """Return characters from the middle of text, raising on Excel errors."""
    return raise_if_sentinel_str(_sentinel_xl_mid(text, start_num, num_chars))


def xl_concatenate(*args: CellValue) -> str:
    """Concatenate text values, raising on Excel errors."""
    return raise_if_sentinel_str(_sentinel_xl_concatenate(*args))


def xl_text(value: CellValue, format_text: CellValue) -> str:
    """Format a value as text, raising on Excel errors."""
    return raise_if_sentinel_str(_sentinel_xl_text(value, format_text))


def xl_numbervalue(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float:
    """Convert text to a number, raising on Excel errors."""
    return raise_if_sentinel_float(
        _sentinel_xl_numbervalue(text, decimal_separator, group_separator)
    )


def xl_lower(text: CellValue) -> str:
    """Return lowercase text, raising on Excel errors."""
    return raise_if_sentinel_str(_sentinel_xl_lower(text))


def xl_value(text: CellValue) -> float:
    """Convert text to a number, preserving Excel ``VALUE`` fallback semantics."""
    parsed = _sentinel_xl_numbervalue(text)
    if parsed is XlError.VALUE and isinstance(text, str):
        return raise_if_sentinel_float(to_number(text))
    return raise_if_sentinel_float(parsed)
