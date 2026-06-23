"""SUMPRODUCT over aligned cell-value arrays (representation-agnostic)."""

from __future__ import annotations

import numpy as np

from .operators_fastpath import try_fastpath_sumproduct
from .operators_reference import reference_sumproduct_arrays
from .types import CellValue, XlError


def xl_sumproduct(*args: CellValue) -> float | XlError:
    if len(args) == 0:
        return 0.0
    arrays: list[np.ndarray] = []
    for arg in args:
        if isinstance(arg, np.ndarray):
            arrays.append(arg)
        else:
            arrays.append(np.array([[arg]], dtype=object))
    shape = arrays[0].shape
    for arr in arrays[1:]:
        if arr.shape != shape:
            return XlError.VALUE

    fast = try_fastpath_sumproduct(arrays)
    if fast is not None:
        return fast
    return reference_sumproduct_arrays(arrays)
