"""Scalar info helpers for exported code without numpy dependencies."""

from __future__ import annotations

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
