"""Raise-only boundary wrappers for reference runtime helpers.

``_sentinel_xl_*`` names are bound at embed time; ``# noqa: F821`` marks those
intentional forward references in wrapper source.
"""

from __future__ import annotations

from excel_grapher.core import CellValue

from .errors import raise_if_sentinel_int, raise_if_sentinel_str

__all__ = ["xl_address", "xl_column", "xl_columns", "xl_row"]


def xl_address(
    row_num: CellValue,
    column_num: CellValue,
    abs_num: CellValue = 1,
    a1: CellValue = True,
    sheet_text: CellValue = None,
) -> str:
    """Build an A1-style address string, raising on Excel errors."""
    return raise_if_sentinel_str(
        _sentinel_xl_address(row_num, column_num, abs_num, a1, sheet_text)  # noqa: F821
    )


def xl_row(ref: CellValue) -> int:
    """Return the row number of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(_sentinel_xl_row(ref))  # noqa: F821


def xl_column(ref: CellValue) -> int:
    """Return the column number of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(_sentinel_xl_column(ref))  # noqa: F821


def xl_columns(ref: CellValue) -> int:
    """Return the column count of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(_sentinel_xl_columns(ref))  # noqa: F821
