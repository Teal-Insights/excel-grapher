from __future__ import annotations

import numpy as np

from ..types import CellValue, XlError
from . import register


@register("ISNUMBER")
def xl_isnumber(value: CellValue) -> bool:
    """Return TRUE if value is a number."""
    if isinstance(value, bool):
        return False  # Booleans are not numbers in Excel's ISNUMBER
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, (np.integer, np.floating)):
        return True
    return False


@register("ISTEXT")
def xl_istext(value: CellValue) -> bool:
    """Return TRUE if value is text."""
    return isinstance(value, str)


@register("ISBLANK")
def xl_isblank(value: CellValue) -> bool:
    """Return TRUE if value is blank (None)."""
    # In Excel, ISBLANK returns TRUE only for empty cells (None)
    # Empty string "" is not considered blank
    return value is None


@register("NA")
def xl_na() -> XlError:
    """Return the #N/A error value."""
    return XlError.NA
