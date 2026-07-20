"""No-op operator fast-path stubs when NumPy is absent or omitted from exports."""

from __future__ import annotations

from typing import Any

from .operator_thresholds import MIN_OPERATOR_FASTPATH_CELLS
from .types import XlError

__all__ = [
    "MIN_OPERATOR_FASTPATH_CELLS",
    "try_fastpath_arithmetic_array",
    "try_fastpath_compare_array",
    "try_fastpath_concat_array",
    "try_fastpath_sumproduct",
]


def try_fastpath_compare_array(
    op: str,
    arr_left: Any,
    arr_right: Any,
) -> Any | XlError | None:
    return None


def try_fastpath_arithmetic_array(
    op: str,
    arr_left: Any,
    arr_right: Any,
) -> Any | XlError | None:
    return None


def try_fastpath_concat_array(
    arr_left: Any,
    arr_right: Any,
) -> Any | None:
    return None


def try_fastpath_sumproduct(arrays: list[Any]) -> float | None:
    return None
