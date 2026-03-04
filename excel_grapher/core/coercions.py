"""Excel-style scalar coercions and value helpers (representation-agnostic)."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from typing import Any

import numpy as np

from .types import CellValue, ExcelRange, XlError


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
