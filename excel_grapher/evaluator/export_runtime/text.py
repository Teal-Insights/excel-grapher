from __future__ import annotations

from .core import CellValue, XlError, to_number, to_string


def xl_left(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    t = to_string(text)
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    return t[: max(0, int(n))]


def xl_right(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    t = to_string(text)
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    k = max(0, int(n))
    return t[-k:] if k else ""


def xl_mid(text: CellValue, start_num: CellValue, num_chars: CellValue) -> str | XlError:
    t = to_string(text)
    s = to_number(start_num)
    if isinstance(s, XlError):
        return s
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    start = int(s) - 1
    length = max(0, int(n))
    if start < 0:
        return ""
    return t[start : start + length]


def xl_concatenate(*args: CellValue) -> str | XlError:
    parts: list[str] = []
    for a in args:
        if isinstance(a, XlError):
            return a
        parts.append(to_string(a))
    return "".join(parts)


def xl_text(value: CellValue, format_text: CellValue) -> str | XlError:
    if isinstance(value, XlError):
        return value
    if isinstance(format_text, XlError):
        return format_text
    return to_string(value)


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

    s = text.replace("\u00A0", " ").strip()
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

