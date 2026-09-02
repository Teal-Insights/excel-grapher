"""Excel operators and series-alignment primitives for inverted-tree codegen.

Mechanical extraction emits calls to these helpers instead of reading cells
from an evaluation context. `take` gathers catalog-order series by index;
internals never see holes and never fetch extra items to pad a result.

A **measure** is a numeric observation or an Excel error code string
(`#REF!`, `#DIV/0!`, …) — the same `err.code` ctx stores on Records. Operators
raise `XlError`; series-member loops catch it so one error cell does not abort
the tuple.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import NoReturn, TypeGuard, TypeVar

T = TypeVar("T")

XL_ERROR_CODES = frozenset(
    {
        "#VALUE!",
        "#REF!",
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NUM!",
        "#NULL!",
    }
)


class XlError(Exception):
    """Excel error value raised as a Python exception."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


def is_error(value: object) -> TypeGuard[str]:
    """True when `value` is an Excel error code string."""
    return isinstance(value, str) and value in XL_ERROR_CODES


def as_measure(value: object, dtype: str = "float") -> int | float | str | bool:
    """Coerce a helper result to a measure: number or cached text.

    Operators still raise `XlError`. Series-member boundaries catch that and
    store `err.code` here so a `#REF!` cell does not abort the rest of a series.
    Non-numeric cached strings (`n/a`, `..`) pass through as measures.
    """
    if isinstance(value, str):
        return value
    if isinstance(value, XlError):
        return value.code
    if dtype == "int":
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            return int(value)
        raise TypeError(f"cannot coerce {type(value).__name__} to int measure")
    if dtype == "str":
        return str(value)
    if dtype == "bool":
        return bool(value)
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, int | float):
        return float(value)
    raise TypeError(f"cannot coerce {type(value).__name__} to float measure")


def xl_div(numerator: float, denominator: float) -> float:
    """Excel `/` with `#DIV/0!` on a zero denominator."""
    if denominator == 0:
        raise XlError("#DIV/0!")
    return numerator / denominator


def xl_choose(index: int, *choices: float) -> float:
    """Excel `CHOOSE`: 1-based selection over already-evaluated arguments."""
    if index < 1 or index > len(choices):
        raise XlError("#VALUE!")
    return choices[index - 1]


def xl_match(lookup: object, lookup_array: Sequence[object], match_type: int = 0) -> int:
    """Excel `MATCH` exact match (`match_type=0`), returning a 1-based index."""
    if match_type != 0:
        raise XlError("#VALUE!")
    for offset, item in enumerate(lookup_array):
        if item == lookup:
            return offset + 1
    raise XlError("#N/A")


def xl_at(values: Sequence[T], index: int) -> T:
    """Return `values[index]` (0-based), raising `#VALUE!` when out of range."""
    if index < 0 or index >= len(values):
        raise XlError("#VALUE!")
    return values[index]


def xl_raise(code: str) -> NoReturn:
    """Raise `XlError(code)` from generated expression position."""
    raise XlError(code)


def require_aligned(*series: Sequence[object]) -> int:
    """Return the common length, or fail if any series has a different length."""
    if not series:
        raise ValueError("require_aligned expected at least one series")
    lengths = [len(item) for item in series]
    if len(set(lengths)) != 1:
        raise ValueError(f"misaligned series lengths: {lengths}")
    return lengths[0]


def require_length(values: Sequence[object], length: int) -> None:
    """Fail if `values` is not a catalog-order array of `length`."""
    actual = len(values)
    if actual != length:
        raise ValueError(f"expected length {length}, got {actual}")


def take(values: Sequence[T], indices: Sequence[int] | slice) -> tuple[T, ...]:
    """Return `values` at 0-based `indices`, failing closed on out-of-range.

    `indices` may be a sequence (including `range`) or a `slice`. A slice
    expands with `range(start, stop, step)`: omitted `start` is 0, omitted
    `stop` is `len(values)`, and `stop` is not clamped — an explicit stop
    past the series length fails closed, same as a tuple of those indices.

    The orchestrator gathers from a catalog-order array into a dense working
    buffer. Internals zip/scan that buffer; they never see holes.
    """
    if isinstance(indices, slice):
        start = 0 if indices.start is None else indices.start
        stop = len(values) if indices.stop is None else indices.stop
        step = 1 if indices.step is None else indices.step
        if step == 0:
            raise ValueError("take slice step cannot be zero")
        indices = range(start, stop, step)
    length = len(values)
    result: list[T] = []
    for index in indices:
        if index < 0 or index >= length:
            raise ValueError(f"take index {index} is outside series of length {length}")
        result.append(values[index])
    return tuple(result)


class InstanceCycleError(ValueError):
    """Demand-driven evaluation hit a same-index circular reference."""


def eval_instance(
    statement: str,
    index: int,
    compute: Callable[[int], T],
    memo: dict[tuple[str, int], T],
    stack: set[tuple[str, int]],
) -> T:
    """Return the memoized value of `statement` at catalog `index`.

    This is the rung-3 dispatcher: demand-driven instance evaluation with an
    on-stack set that raises `InstanceCycleError` on a real cycle.
    """
    if index < 0:
        raise XlError("#REF!")
    key = (statement, index)
    if key in memo:
        return memo[key]
    if key in stack:
        raise InstanceCycleError(f"distance-zero cycle evaluating {statement}[{index}]")
    stack.add(key)
    try:
        value = compute(index)
    finally:
        stack.remove(key)
    memo[key] = value
    return value


def live_measure(value: T) -> T:
    """Return `value`, or raise `XlError` when it is a stored error code."""
    if isinstance(value, str) and is_error(value):
        raise XlError(value)
    return value


def demand_instance(
    statement: str,
    index: int,
    compute: Callable[[int], T],
    memo: dict[tuple[str, int], T],
    stack: set[tuple[str, int]],
) -> T:
    """Like `eval_instance`, but re-raise a stored Excel error as `XlError`."""
    value = eval_instance(statement, index, compute, memo, stack)
    if isinstance(value, str) and is_error(value):
        raise XlError(value)
    return value
