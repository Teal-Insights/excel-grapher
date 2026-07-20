"""Raise-only boundary wrapper for SUMPRODUCT over shared Grid traversal."""

from __future__ import annotations

from excel_grapher.core import CellValue
from excel_grapher.core.sumproduct import sumproduct_cells

from .errors import raise_if_sentinel_float

__all__ = ["xl_sumproduct"]


def xl_sumproduct(*args: CellValue) -> float:
    """Return the sum of element-wise products, raising on Excel errors."""
    return raise_if_sentinel_float(sumproduct_cells(*args))
