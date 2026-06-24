"""No-op operator fast-path stubs for exports that omit vectorized code."""

from __future__ import annotations

import numpy as np

from .types import XlError


def try_fastpath_compare_array(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | XlError | None:
    return None


def try_fastpath_arithmetic_array(
    op: str,
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | XlError | None:
    return None


def try_fastpath_concat_array(
    arr_left: np.ndarray,
    arr_right: np.ndarray,
) -> np.ndarray | None:
    return None


def try_fastpath_sumproduct(arrays: list[np.ndarray]) -> float | None:
    return None
