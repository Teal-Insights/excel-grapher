"""Excel operators and series-alignment primitives for inverted-tree codegen.

Mechanical extraction emits calls to these helpers instead of reading cells
from an evaluation context. `take` gathers catalog-order series by index;
internals never see holes and never fetch extra items to pad a result.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import NoReturn, TypeVar

T = TypeVar("T")


class XlError(Exception):
    """Excel error value raised as a Python exception."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


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


def take(values: Sequence[T], indices: Sequence[int]) -> tuple[T, ...]:
    """Return `values` at 0-based `indices`, failing closed on out-of-range.

    The orchestrator gathers from a catalog-order array into a dense working
    buffer. Internals zip/scan that buffer; they never see holes.
    """
    length = len(values)
    result: list[T] = []
    for index in indices:
        if index < 0 or index >= length:
            raise ValueError(f"take index {index} is outside series of length {length}")
        result.append(values[index])
    return tuple(result)
