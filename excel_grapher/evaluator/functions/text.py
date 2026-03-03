from __future__ import annotations

from ..helpers import to_number, to_string
from ..export_runtime.text import xl__xlfn_numbervalue as _rt_xl__xlfn_numbervalue
from ..types import CellValue, XlError
from . import register


@register("LEFT")
def xl_left(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    """Return leftmost characters from a text string."""
    s = to_string(text)
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    chars = int(n)
    if chars < 0:
        return XlError.VALUE
    return s[:chars]


@register("RIGHT")
def xl_right(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    """Return rightmost characters from a text string."""
    s = to_string(text)
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    chars = int(n)
    if chars < 0:
        return XlError.VALUE
    if chars == 0:
        return ""
    return s[-chars:]


@register("MID")
def xl_mid(
    text: CellValue, start_num: CellValue, num_chars: CellValue
) -> str | XlError:
    """Return characters from the middle of a text string."""
    s = to_string(text)
    start = to_number(start_num)
    if isinstance(start, XlError):
        return start
    num = to_number(num_chars)
    if isinstance(num, XlError):
        return num
    start_idx = int(start) - 1  # Excel uses 1-based indexing
    chars = int(num)
    if start_idx < 0 or chars < 0:
        return XlError.VALUE
    return s[start_idx : start_idx + chars]


@register("CONCATENATE")
def xl_concatenate(*args: CellValue) -> str:
    """Join several text strings into one."""
    return "".join(to_string(arg) for arg in args)


@register("TEXT")
def xl_text(value: CellValue, format_text: CellValue) -> str | XlError:
    """Format a number as text with a specified format.

    Note: This is a simplified implementation supporting common formats.
    """
    fmt = to_string(format_text)
    n = to_number(value)
    if isinstance(n, XlError):
        # If not a number, just return string representation
        return to_string(value)

    # Simple format handling
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

    # Default: return general number format
    if n == int(n):
        return str(int(n))
    return str(n)


@register("_XLFN.NUMBERVALUE")
@register("NUMBERVALUE")
def xl__xlfn_numbervalue(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float | XlError:
    """Convert text to number with optional separators."""
    return _rt_xl__xlfn_numbervalue(text, decimal_separator, group_separator)
