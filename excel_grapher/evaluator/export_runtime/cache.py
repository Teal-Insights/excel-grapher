from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
import warnings

import openpyxl.utils.cell

from .core import CellValue, ExcelRange, XlError

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


@dataclass(slots=True)
class EvalContext:
    """Per-run evaluation state for generated spreadsheets.

    The exported-code path needs a mutable inputs mapping and a cache that is scoped
    to a single compute call, so callers can run many scenarios without global state.
    """

    inputs: dict[str, CellValue]
    resolver: Callable[[str], Callable[[EvalContext], CellValue] | None]
    cache: dict[str, CellValue] = field(default_factory=dict)
    computing: set[str] = field(default_factory=set)
    deps: dict[str, set[str]] = field(default_factory=dict)
    reverse_deps: dict[str, set[str]] = field(default_factory=dict)
    stack: list[str] = field(default_factory=list)

    def _record_dependency(self, parent: str, child: str) -> None:
        if parent == child:
            return
        self.deps.setdefault(parent, set()).add(child)
        self.reverse_deps.setdefault(child, set()).add(parent)

    def invalidate(self, addresses: Iterable[str]) -> None:
        """Invalidate cached values for the given addresses and their dependents."""
        to_visit = list(addresses)
        seen: set[str] = set()
        while to_visit:
            addr = to_visit.pop()
            if addr in seen:
                continue
            seen.add(addr)

            self.cache.pop(addr, None)
            self.computing.discard(addr)

            dependents = list(self.reverse_deps.get(addr, set()))
            to_visit.extend(dependents)

            for dep in self.deps.get(addr, set()):
                parents = self.reverse_deps.get(dep)
                if parents is not None:
                    parents.discard(addr)
                    if not parents:
                        self.reverse_deps.pop(dep, None)

            self.deps.pop(addr, None)
            self.reverse_deps.pop(addr, None)

    def set_inputs(self, inputs: dict[str, CellValue]) -> None:
        """Update input values and invalidate dependent cached results."""
        changed = [k for k, v in inputs.items() if self.inputs.get(k) != v]
        self.inputs.update(inputs)
        if changed:
            self.invalidate(changed)


def xl_cell(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a single cell address under the given context.

    Resolution order:
    - cached value (per ctx)
    - user-provided inputs
    - exported formula implementation (via resolver)
    - missing cell raises KeyError
    """
    if ctx.stack:
        ctx._record_dependency(ctx.stack[-1], address)

    if address in ctx.cache:
        return ctx.cache[address]

    if address in ctx.computing:
        return xl_circular_reference()

    if address in ctx.inputs:
        v = ctx.inputs[address]
        ctx.cache[address] = v
        return v

    fn = ctx.resolver(address)
    if fn is None:
        raise KeyError(f"Cell {address} not found in graph")

    ctx.computing.add(address)
    ctx.stack.append(address)
    try:
        v = fn(ctx)
        # Excel treats "empty" results as 0 in most numeric contexts; the evaluator
        # normalizes formula results of None to 0, so do the same here for parity.
        if v is None:
            v = 0
        ctx.cache[address] = v
        return v
    finally:
        ctx.computing.discard(address)
        if ctx.stack and ctx.stack[-1] == address:
            ctx.stack.pop()


def xl_eval(
    ctx: EvalContext,
    address: str,
    fn: Callable[[EvalContext], CellValue],
) -> CellValue:
    """Evaluate a known formula implementation under the given context."""
    if ctx.stack:
        ctx._record_dependency(ctx.stack[-1], address)

    if address in ctx.cache:
        return ctx.cache[address]

    if address in ctx.computing:
        return xl_circular_reference()

    if address in ctx.inputs:
        v = ctx.inputs[address]
        ctx.cache[address] = v
        return v

    ctx.computing.add(address)
    ctx.stack.append(address)
    try:
        v = fn(ctx)
        if v is None:
            v = 0
        ctx.cache[address] = v
        return v
    finally:
        ctx.computing.discard(address)
        if ctx.stack and ctx.stack[-1] == address:
            ctx.stack.pop()


def _parse_sheet_address(address: str) -> tuple[str, str] | None:
    if address.startswith("'"):
        i = 1
        while i < len(address):
            if address[i] == "'":
                if i + 1 < len(address) and address[i + 1] == "'":
                    i += 2
                    continue
                break
            i += 1
        sheet = address[: i + 1]
        rest = address[i + 1 :]
        if rest.startswith("!"):
            return sheet, rest[1:]
        return None

    if "!" in address:
        sheet, cell = address.rsplit("!", 1)
        return sheet, cell

    return None


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
        start_col, start_row = openpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = openpyxl.utils.cell.coordinate_from_string(end_cell)
        start_col_idx = openpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = openpyxl.utils.cell.column_index_from_string(end_col)
    except ValueError:
        return XlError.VALUE

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx

    rng = ExcelRange(sheet, start_row, start_col_idx, end_row, end_col_idx)
    return rng.resolve(lambda addr: xl_cell(ctx, addr))
