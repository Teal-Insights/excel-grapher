"""Slim cache eval helpers for exports that omit dependency tracking."""

from __future__ import annotations

from collections.abc import Callable
from typing import cast

from excel_grapher.core import CellValue, XlError, XlErrorException

from .cache import xl_circular_reference
from .cache_context import EvalContextBase
from .leaves import MISSING, lookup_leaf

__all__ = [
    "EvalContext",
    "_evaluate_address",
    "xl_cell",
    "xl_eval",
]

EvalContext = EvalContextBase


def _raise_if_error_value(value: CellValue) -> CellValue:
    """Surface Excel error values as raised exceptions at the cell boundary."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value


def _evaluate_address(
    ctx: EvalContextBase,
    address: str,
    obtain_fn: Callable[[], Callable[[EvalContextBase], CellValue]],
    *,
    preserve_structural_blank: bool = False,
) -> CellValue:
    """Shared evaluation path without dependency recording (slim export scaffold).

    Excel error values raise `XlErrorException`; the raising cell's error code
    is cached so re-reads raise without re-evaluating.
    """
    if address in ctx.cache:
        return _raise_if_error_value(ctx.cache[address])

    if address in ctx.computing:
        if ctx.iterative_enabled:
            return ctx.iteration_values.get(address, 0)
        return xl_circular_reference()

    found = lookup_leaf(ctx, address)
    if found is not MISSING:
        ctx.cache[address] = found
        return _raise_if_error_value(found)

    fn = obtain_fn()

    ctx.computing.add(address)
    try:
        try:
            value = fn(ctx)
        except XlErrorException as exc:
            ctx.cache[address] = exc.code
            raise
        if value is None and not (
            preserve_structural_blank and getattr(fn, "__structural_blank__", False)
        ):
            value = 0
        ctx.cache[address] = value
        return _raise_if_error_value(value)
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
