"""Helpers for top-level array formula results (issue #284)."""

from __future__ import annotations

from typing import cast

import numpy as np

from excel_grapher.core.types import CellValue, XlError


def is_array_result(value: CellValue) -> bool:
    """Return whether ``value`` is a multi-cell array formula result."""
    return isinstance(value, np.ndarray) and value.size > 1


def array_values_equal(a: object, b: object) -> bool:
    """Compare two values, including element-wise for matching ``ndarray`` shapes."""
    if isinstance(a, np.ndarray) and isinstance(b, np.ndarray):
        left = cast(np.ndarray, a)
        right = cast(np.ndarray, b)
        if left.shape != right.shape:
            return False
        for index in range(int(left.size)):
            if not _scalar_values_equal(left.flat[index], right.flat[index]):
                return False
        return True
    return _scalar_values_equal(a, b)


def _scalar_values_equal(a: object, b: object) -> bool:
    if a is b:
        return True
    if isinstance(a, XlError) or isinstance(b, XlError):
        return a == b
    if isinstance(a, (bool, np.bool_)) and isinstance(b, (bool, np.bool_)):
        return bool(a) == bool(b)
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return float(a) == float(b)
    return a == b
