"""Excel-style scalar coercions and value helpers (representation-agnostic)."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from datetime import date, datetime
from typing import Any

import numpy as np

from .types import CellValue, ExcelRange, XlError

_EXCEL_EPOCH = datetime(1899, 12, 30)


def datetime_to_excel_serial(value: datetime) -> float:
    """Convert a naive datetime to an Excel day serial (1900 date system)."""
    naive = value.replace(tzinfo=None) if value.tzinfo is not None else value
    delta = naive - _EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86_400.0


def _try_parse_iso_date_serial(text: str) -> float | None:
    stripped = text.strip()
    if not stripped:
        return None
    try:
        if "T" in stripped or " " in stripped:
            parsed = datetime.fromisoformat(stripped.replace("Z", "+00:00"))
            if parsed.tzinfo is not None:
                parsed = parsed.replace(tzinfo=None)
        else:
            parsed = datetime.combine(date.fromisoformat(stripped), datetime.min.time())
        return datetime_to_excel_serial(parsed)
    except ValueError:
        return None


def to_native(value: Any) -> Any:
    if hasattr(value, "item"):
        return value.item()
    return value


def to_number(value: CellValue) -> float | XlError:
    if value is None:
        return 0.0
    if isinstance(value, XlError):
        return value
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if s == "":
            return 0.0
        try:
            return float(s)
        except ValueError:
            serial = _try_parse_iso_date_serial(s)
            if serial is not None:
                return serial
            return XlError.VALUE
    if isinstance(value, ExcelRange):
        return XlError.VALUE
    return XlError.VALUE


def to_int(value: CellValue) -> int | XlError:
    """Coerce a CellValue to an integer using Excel-style numeric coercion.

    For functions that operate on integer indices (e.g. CHOOSE/INDEX/MATCH)
    while propagating Excel errors.
    """
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return int(n)


def _format_general_number(value: float | int) -> str:
    f = float(value)
    if f.is_integer():
        return str(int(f))
    return str(f)


def to_string(value: CellValue) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, XlError):
        return value.value
    if isinstance(value, (int, float)):
        return _format_general_number(float(value))
    if isinstance(value, ExcelRange):
        return XlError.VALUE.value
    return str(value)


def to_bool(value: CellValue) -> bool | XlError:
    if value is None:
        return False
    if isinstance(value, XlError):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0
    if isinstance(value, str):
        s = value.strip().upper()
        if s == "":
            return False
        if s == "TRUE":
            return True
        if s == "FALSE":
            return False
        return XlError.VALUE
    if isinstance(value, ExcelRange):
        return XlError.VALUE
    return XlError.VALUE


def excel_casefold(value: str) -> str:
    return value.casefold()


def flatten(*args: Any) -> Iterator[CellValue]:
    for arg in args:
        if isinstance(arg, np.ndarray):
            yield from (v for v in arg.flat)
            continue
        if isinstance(arg, (list, tuple)):
            yield from flatten(*arg)
            continue
        yield arg


def get_error(*args: Any) -> XlError | None:
    for v in flatten(*args):
        if isinstance(v, XlError):
            return v
    return None


def numeric_values(values: Iterable[CellValue]) -> tuple[list[float], XlError | None]:
    nums: list[float] = []
    for v in values:
        n = to_number(v)
        if isinstance(n, XlError):
            return ([], n)
        nums.append(float(n))
    return (nums, None)
