from __future__ import annotations

import warnings
from collections.abc import Callable, Mapping
from math import isfinite
from typing import cast

import fastpyxl.utils.cell

from excel_grapher.core import CellValue, ExcelRange, XlError
from excel_grapher.core.address_keys import normalize_key
from excel_grapher.core.addressing import split_sheet_qualified_address
from excel_grapher.core.array_results import (
    finalize_top_level_array_result,
    read_spill_scalar,
    scalar_for_range_member,
)

from .cache_context import EvalContext, EvalContextBase

__all__ = [
    "CircularReferenceWarning",
    "EvalContext",
    "EvalContextBase",
    "circular_safe_cache",
    "coerce_inputs_dict",
    "xl_cell",
    "xl_cell_in_range",
    "xl_circular_reference",
    "xl_eval",
    "xl_iterative_compute",
    "xl_range",
]

_cell_cache: dict[Callable[[], CellValue], CellValue] = {}
_computing: set[Callable[[], CellValue]] = set()


class CircularReferenceWarning(RuntimeWarning):
    """Warning emitted when a circular reference is encountered (default Excel mode)."""


def xl_circular_reference() -> CellValue:
    """Excel default behavior for circular references (non-iterative calculation)."""
    warnings.warn(
        "Circular reference detected; returning 0 (iterative calculation is disabled).",
        CircularReferenceWarning,
        stacklevel=2,
    )
    return 0


def circular_safe_cache(func: Callable[[], CellValue]) -> Callable[[], CellValue]:
    """Cache decorator that breaks circular references by returning 0."""

    def wrapper() -> CellValue:
        if func in _computing:
            return xl_circular_reference()
        if func in _cell_cache:
            return _cell_cache[func]
        _computing.add(func)
        try:
            result = func()
            _cell_cache[func] = result
            return result
        finally:
            _computing.discard(func)

    return wrapper


def coerce_inputs_dict(values: Mapping[str, object]) -> dict[str, CellValue]:
    """Widen inferred default-input dicts to `dict[str, CellValue]` for `EvalContext`."""
    return cast(dict[str, CellValue], dict(values))


def _evaluate_address(
    ctx: EvalContext,
    address: str,
    obtain_fn: Callable[[], Callable[[EvalContext], CellValue]],
    *,
    preserve_structural_blank: bool = False,
) -> CellValue:
    """Shared evaluation path for ``xl_cell`` and ``xl_eval``."""
    if ctx.stack:
        ctx._record_dependency(ctx.stack[-1], address)

    if address in ctx.cache:
        return ctx.cache[address]

    if address in ctx.computing:
        if ctx.iterative_enabled:
            return ctx.iteration_values.get(address, 0)
        return xl_circular_reference()

    if address in ctx.inputs:
        v = ctx.inputs[address]
        ctx.cache[address] = v
        return v

    fn = obtain_fn()

    ctx.computing.add(address)
    ctx.stack.append(address)
    try:
        v = fn(ctx)
        if v is None and not (
            preserve_structural_blank and getattr(fn, "__structural_blank__", False)
        ):
            v = 0
        occupant = ctx.spill_is_occupied
        if occupant is not None:
            v = finalize_top_level_array_result(address, v, is_occupied=occupant)
        ctx.cache[address] = v
        return v
    finally:
        ctx.computing.discard(address)
        if ctx.stack and ctx.stack[-1] == address:
            ctx.stack.pop()


def xl_cell_in_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate one range member, projecting array anchors to spilled scalars."""
    norm = normalize_key(address)
    if ctx.resolver(norm) is not None or norm in ctx.inputs or norm in ctx.cache:
        value = xl_cell(ctx, norm)
        return scalar_for_range_member(norm, value)
    spilled = read_spill_scalar(norm, ctx.cache)
    if spilled is not None:
        return spilled
    raise KeyError(f"Cell {address} not found in graph")


def xl_cell(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a single cell address under the given context.

    Resolution order:
    - cached value (per ctx)
    - user-provided inputs
    - exported formula implementation (via resolver)
    - missing cell raises KeyError
    """

    def obtain_fn() -> Callable[[EvalContext], CellValue]:
        fn = ctx.resolver(address)
        if fn is None:
            raise KeyError(f"Cell {address} not found in graph")
        return fn

    return _evaluate_address(ctx, address, obtain_fn, preserve_structural_blank=True)


def xl_eval(
    ctx: EvalContext,
    address: str,
    fn: Callable[[EvalContext], CellValue],
) -> CellValue:
    """Evaluate a known formula implementation under the given context."""
    return _evaluate_address(ctx, address, lambda: fn, preserve_structural_blank=False)


def _parse_sheet_address(address: str) -> tuple[str, str] | None:
    return split_sheet_qualified_address(address)


def _parse_range_address(address: str) -> tuple[str, str, str] | XlError:
    if ":" not in address:
        return XlError.VALUE
    start_text, end_text = address.split(":", 1)
    start = _parse_sheet_address(start_text)
    if start is None:
        return XlError.VALUE
    sheet, start_cell = start
    if "!" in end_text:
        end = _parse_sheet_address(end_text)
        if end is None:
            return XlError.VALUE
        end_sheet, end_cell = end
        if end_sheet != sheet:
            return XlError.VALUE
    else:
        end_cell = end_text
    return sheet, start_cell, end_cell


def xl_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range and return a 2D numpy array of values."""
    parsed = _parse_range_address(address)
    if isinstance(parsed, XlError):
        return parsed
    sheet, start_cell, end_cell = parsed
    try:
        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)
        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)
    except ValueError:
        return XlError.VALUE

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx

    rng = ExcelRange(sheet, start_row, start_col_idx, end_row, end_col_idx)
    return rng.resolve(lambda addr: xl_cell_in_range(ctx, addr))


def _convergence_delta(prev: CellValue, curr: CellValue) -> float:
    if hasattr(prev, "shape") and hasattr(curr, "shape"):
        try:
            from typing import Any

            import numpy as np

            def _is_array(v: Any) -> bool:
                return hasattr(v, "shape")

            if _is_array(prev) and _is_array(curr):
                from typing import cast

                return (
                    0.0
                    if np.array_equal(cast(Any, prev), cast(Any, curr), equal_nan=True)
                    else float("inf")
                )
            return float("inf")
        except Exception:
            return float("inf")

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


def _has_converged(
    previous: dict[str, CellValue],
    current: dict[str, CellValue],
    *,
    iterate_delta: float,
) -> bool:
    for key, curr in current.items():
        prev = previous.get(key, 0)
        if _convergence_delta(prev, curr) > iterate_delta:
            return False
    return True


def xl_iterative_compute(
    ctx: EvalContext,
    targets: dict[str, Callable[[EvalContext, str], CellValue]],
) -> dict[str, CellValue]:
    """Compute targets with Excel-style iterative convergence semantics."""
    previous = dict(ctx.iteration_values)
    iterations = max(1, int(ctx.iterate_count))
    for _ in range(iterations):
        ctx.cache.clear()
        ctx.computing.clear()
        ctx.stack.clear()
        ctx.iteration_values.clear()
        ctx.iteration_values.update(previous)
        current = {target: handler(ctx, target) for target, handler in targets.items()}
        next_previous = dict(current)
        next_previous.update(ctx.cache)
        if _has_converged(previous, next_previous, iterate_delta=float(ctx.iterate_delta)):
            ctx.iteration_values.clear()
            ctx.iteration_values.update(next_previous)
            return current
        previous = next_previous

    ctx.iteration_values.clear()
    ctx.iteration_values.update(previous)
    return {target: handler(ctx, target) for target, handler in targets.items()}
