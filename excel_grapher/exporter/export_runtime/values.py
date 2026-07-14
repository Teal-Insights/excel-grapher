"""Export-owned Excel value model shared by exported runtime modules."""

from __future__ import annotations

from collections.abc import Iterator
from math import isfinite
from typing import TypeAlias, cast

from excel_grapher.core import XlError
from excel_grapher.core.grid import Grid, Range, Scalar
from excel_grapher.core.types import ExcelRange, XlErrorException

__all__ = ["CellValue", "ExcelRange", "Grid", "Scalar", "as_scalar", "flatten"]


CellValue: TypeAlias = Scalar | ExcelRange | Range | list["CellValue"]


def as_scalar(value: CellValue) -> Scalar:
    """Collapse range/array values to `#VALUE!` for scalar coercion contexts.

    Keep behavior aligned with `excel_grapher.core.coercions.as_scalar`. This
    module is embedded into standalone exports and cannot import library code.
    """
    if isinstance(value, (Range, ExcelRange, list, tuple)):
        return XlError.VALUE
    return value


def _convergence_delta(prev: CellValue, curr: CellValue) -> float:
    if isinstance(prev, (Range, list)) or isinstance(curr, (Range, list)):
        prev_rows = prev.rows_raw() if isinstance(prev, Range) else prev
        curr_rows = curr.rows_raw() if isinstance(curr, Range) else curr
        return 0.0 if prev_rows == curr_rows else float("inf")

    if isinstance(prev, bool) or isinstance(curr, bool):
        return 0.0 if prev == curr else float("inf")
    if isinstance(prev, (int, float)) and isinstance(curr, (int, float)):
        pf = float(prev)
        cf = float(curr)
        if isfinite(pf) and isfinite(cf):
            return abs(cf - pf)
    try:
        eq = prev == curr
    except Exception:
        return float("inf")
    if isinstance(eq, bool):
        return 0.0 if eq else float("inf")
    return float("inf")


def flatten(*args: CellValue) -> Iterator[Scalar]:
    """Yield scalar values from scalars, nested lists, and lazy ranges.

    Excel errors raise `XlErrorException` on encounter, mirroring the
    evaluator's argument error precheck for generic worksheet functions.
    Lookup scans keep skip semantics through `Grid`/`Range.value_at` instead.
    """
    for arg in args:
        if isinstance(arg, Range):
            for v in arg.iter_raw():
                if isinstance(v, XlError):
                    raise XlErrorException(v)
                yield cast(Scalar, v)
        elif isinstance(arg, (list, tuple)):
            yield from flatten(*arg)
        else:
            if isinstance(arg, XlError):
                raise XlErrorException(arg)
            yield cast(Scalar, arg)
