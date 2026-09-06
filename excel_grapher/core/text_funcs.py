"""Shared text worksheet function implementations."""

from __future__ import annotations

from typing import cast

from .coercions import as_scalar, to_number, to_string
from .types import CellValue, XlError

__all__ = [
    "concatenate_cells",
    "left_chars",
    "lower_text",
    "mid_chars",
    "numbervalue_parse",
    "right_chars",
    "text_format",
    "value_from_text",
]


def left_chars(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    """Return the leftmost characters of text."""
    scalar = as_scalar(text)
    if isinstance(scalar, XlError):
        return scalar
    s = to_string(cast(CellValue, scalar))
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    chars = int(n)
    if chars < 0:
        return XlError.VALUE
    return s[:chars]


def right_chars(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    """Return the rightmost characters of text."""
    scalar = as_scalar(text)
    if isinstance(scalar, XlError):
        return scalar
    s = to_string(cast(CellValue, scalar))
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    chars = int(n)
    if chars < 0:
        return XlError.VALUE
    if chars == 0:
        return ""
    return s[-chars:]


def mid_chars(text: CellValue, start_num: CellValue, num_chars: CellValue) -> str | XlError:
    """Return characters from the middle of text."""
    scalar = as_scalar(text)
    if isinstance(scalar, XlError):
        return scalar
    s = to_string(cast(CellValue, scalar))
    start = to_number(start_num)
    if isinstance(start, XlError):
        return start
    num = to_number(num_chars)
    if isinstance(num, XlError):
        return num
    start_idx = int(start) - 1
    chars = int(num)
    if start_idx < 0 or chars < 0:
        return XlError.VALUE
    return s[start_idx : start_idx + chars]


def concatenate_cells(*args: CellValue) -> str | XlError:
    """Concatenate text values."""
    parts: list[str] = []
    for a in args:
        if isinstance(a, XlError):
            return a
        scalar = as_scalar(a)
        if isinstance(scalar, XlError):
            return scalar
        parts.append(to_string(cast(CellValue, scalar)))
    return "".join(parts)


def text_format(value: CellValue, format_text: CellValue) -> str | XlError:
    """Format a value as text using a format string."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        return scalar
    if isinstance(format_text, XlError):
        return format_text
    fmt = to_string(format_text)
    n = to_number(cast(CellValue, scalar))
    if isinstance(n, XlError):
        return to_string(cast(CellValue, scalar))

    if fmt == "0":
        return str(int(round(n)))
    if fmt == "0.0":
        return f"{n:.1f}"
    if fmt == "0.00":
        return f"{n:.2f}"
    if fmt == "0.000":
        return f"{n:.3f}"
    if fmt == "#,##0":
        return f"{int(round(n)):,}"
    if fmt == "#,##0.00":
        return f"{n:,.2f}"
    if fmt == "0%":
        return f"{int(round(n * 100))}%"
    if fmt == "0.0%":
        return f"{n * 100:.1f}%"
    if fmt == "0.00%":
        return f"{n * 100:.2f}%"

    if n == int(n):
        return str(int(n))
    return str(n)


def numbervalue_parse(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float | XlError:
    """Convert text to a number with explicit decimal and group separators."""
    if isinstance(text, XlError):
        return text
    if isinstance(decimal_separator, XlError):
        return decimal_separator
    if isinstance(group_separator, XlError):
        return group_separator

    if not isinstance(text, str):
        return to_number(text)

    dec_sep = to_string(decimal_separator)
    grp_sep = to_string(group_separator)
    if dec_sep == "" or dec_sep == grp_sep:
        return XlError.VALUE

    s = text.replace("\u00a0", " ").strip()
    if s == "":
        return 0.0
    currency_symbols = "$€£¥"
    while s and (s[0] in currency_symbols or s[-1] in currency_symbols):
        s = s.lstrip(currency_symbols).rstrip(currency_symbols).strip()
        if s == "":
            return XlError.VALUE
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()
        if s == "":
            return XlError.VALUE
    percent = False
    if s.endswith("%"):
        percent = True
        s = s[:-1].strip()
        if s == "":
            return XlError.VALUE
    sign = 1.0
    if s.startswith(("+", "-")):
        if s[0] == "-":
            sign = -1.0
        s = s[1:].strip()
        if s == "":
            return XlError.VALUE
    while s and (s[0] in currency_symbols or s[-1] in currency_symbols):
        s = s.lstrip(currency_symbols).rstrip(currency_symbols).strip()
        if s == "":
            return XlError.VALUE
    if grp_sep:
        s = s.replace(grp_sep, "")
    if dec_sep != ".":
        s = s.replace(dec_sep, ".")
    try:
        value = float(s)
    except ValueError:
        return XlError.VALUE
    if percent:
        value /= 100.0
    if negative:
        value = -abs(value)
    return value * sign


def lower_text(text: CellValue) -> str | XlError:
    """Return lowercase text."""
    if isinstance(text, XlError):
        return text
    return to_string(text).lower()


def value_from_text(text: CellValue) -> float | XlError:
    """Convert locale-formatted text to a number (Excel VALUE)."""
    parsed = numbervalue_parse(text)
    if parsed is XlError.VALUE and isinstance(text, str):
        return to_number(text)
    return parsed
