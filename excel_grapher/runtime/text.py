from __future__ import annotations

from excel_grapher.core import CellValue, XlError
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


def xl_left(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    return left_chars(text, num_chars)


def xl_right(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    return right_chars(text, num_chars)


def xl_mid(text: CellValue, start_num: CellValue, num_chars: CellValue) -> str | XlError:
    return mid_chars(text, start_num, num_chars)


def xl_concatenate(*args: CellValue) -> str | XlError:
    return concatenate_cells(*args)


def xl_text(value: CellValue, format_text: CellValue) -> str | XlError:
    return text_format(value, format_text)


def xl_numbervalue(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float | XlError:
    """Convert text to a number with explicit decimal and group separators."""
    return numbervalue_parse(text, decimal_separator, group_separator)


def xl_lower(text: CellValue) -> str | XlError:
    return lower_text(text)


def xl_value(text: CellValue) -> float | XlError:
    """Convert locale-formatted text to a number (Excel VALUE)."""
    return value_from_text(text)
