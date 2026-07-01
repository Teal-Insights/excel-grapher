"""Scalar info helpers for exported code without numpy dependencies."""

from __future__ import annotations

from excel_grapher.core import XlError

from .values import CellValue, flatten

__all__ = ["xl_count", "xl_isnumber"]


def xl_isnumber(value: CellValue) -> bool:
    return not isinstance(value, bool) and isinstance(value, (int, float))


def xl_count(*args: CellValue) -> int:
    count = 0
    for v in flatten(*args):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            count += 1
    return count


def _iter_numeric_cells(values: list[CellValue]) -> tuple[list[float], XlError | None]:
    nums: list[float] = []
    for v in values:
        if isinstance(v, XlError):
            return ([], v)
        if v is None:
            continue
        if isinstance(v, bool):
            continue
        if isinstance(v, (int, float)):
            nums.append(float(v))
    return (nums, None)
