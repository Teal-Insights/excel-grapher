"""Slim cache eval helpers for exports with top-level array / spill semantics."""

from __future__ import annotations

from collections.abc import Callable
from typing import cast

from excel_grapher.core import CellValue
from excel_grapher.core.array_results import finalize_top_level_array_result

from .cache import xl_circular_reference
from .cache_context import EvalContextBase

__all__ = [
    "EvalContext",
    "_evaluate_address",
    "xl_cell",
    "xl_eval",
]

EvalContext = EvalContextBase


def _evaluate_address(
    ctx: EvalContextBase,
    address: str,
    obtain_fn: Callable[[], Callable[[EvalContextBase], CellValue]],
    *,
    preserve_structural_blank: bool = False,
) -> CellValue:
    """Shared evaluation path without dependency recording (slim export scaffold)."""
    if address in ctx.cache:
        return ctx.cache[address]

    if address in ctx.computing:
        if ctx.iterative_enabled:
            return ctx.iteration_values.get(address, 0)
        return xl_circular_reference()

    if address in ctx.inputs:
        value = ctx.inputs[address]
        ctx.cache[address] = value
        return value

    fn = obtain_fn()

    ctx.computing.add(address)
    try:
        value = fn(ctx)
        if value is None and not (
            preserve_structural_blank and getattr(fn, "__structural_blank__", False)
        ):
            value = 0
        occupant = ctx.spill_is_occupied
        if occupant is not None:
            value = finalize_top_level_array_result(address, value, is_occupied=occupant)
        ctx.cache[address] = value
        return value
    finally:
        ctx.computing.discard(address)


def xl_cell(ctx: EvalContextBase, address: str) -> CellValue:
    """Evaluate a single cell address under the given context (slim scaffold)."""

    def obtain_fn() -> Callable[[EvalContextBase], CellValue]:
        fn = ctx.resolver(address)
        if fn is None:
            raise KeyError(f"Cell {address} not found in graph")
        return cast(Callable[[EvalContextBase], CellValue], fn)

    return _evaluate_address(ctx, address, obtain_fn, preserve_structural_blank=True)


def xl_eval(
    ctx: EvalContextBase,
    address: str,
    fn: Callable[[EvalContextBase], CellValue],
) -> CellValue:
    """Evaluate a known formula implementation under the given context (slim scaffold)."""
    return _evaluate_address(ctx, address, lambda: fn, preserve_structural_blank=False)
