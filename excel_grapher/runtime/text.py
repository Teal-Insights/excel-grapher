from __future__ import annotations

from excel_grapher.core import CellValue, XlError, to_number, to_string

__all__ = [
    "xl__xlfn_numbervalue",
    "xl_concatenate",
    "xl_left",
    "xl_mid",
    "xl_numbervalue",
    "xl_right",
    "xl_text",
]


def xl_left(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    s = to_string(text)
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    chars = int(n)
    if chars < 0:
        return XlError.VALUE
    return s[:chars]


def xl_right(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
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


def xl_mid(text: CellValue, start_num: CellValue, num_chars: CellValue) -> str | XlError:
    s = to_string(text)
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


def xl_concatenate(*args: CellValue) -> str | XlError:
    parts: list[str] = []
    for a in args:
        if isinstance(a, XlError):
            return a
        parts.append(to_string(a))
    return "".join(parts)


def xl_text(value: CellValue, format_text: CellValue) -> str | XlError:
    fmt = to_string(format_text)
    n = to_number(value)
    if isinstance(n, XlError):
        return to_string(value)

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


def xl__xlfn_numbervalue(
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


def xl_numbervalue(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float | XlError:
    """Excel NUMBERVALUE wrapper (non _xlfn prefix)."""
    return xl__xlfn_numbervalue(text, decimal_separator, group_separator)
