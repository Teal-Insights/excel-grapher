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

import math
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


def _as_number(value: object) -> float:
    """Coerce `value` to a float the way Excel numeric operators do.

    Blank (`None`) and empty / whitespace-only text are `0`. Booleans are
    `1` / `0`. A stored error code is re-raised; any other non-numeric string
    is `#VALUE!`.
    """
    if isinstance(value, str) and is_error(value):
        raise XlError(value)
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, int | float):
        return float(value)
    if isinstance(value, str):
        text = value.replace("\u00a0", "").replace(" ", "")
        if text == "":
            return 0.0
        try:
            return float(text)
        except ValueError:
            raise XlError("#VALUE!") from None
    raise XlError("#VALUE!")


def _try_number(value: object) -> float | None:
    """Return `_as_number(value)`, or `None` when coercion is `#VALUE!`.

    Stored error-code measures (including `#VALUE!`) are re-raised so
    comparisons do not treat a cached error as failed numeric coercion.
    """
    if isinstance(value, str) and is_error(value):
        raise XlError(value)
    try:
        return _as_number(value)
    except XlError as err:
        if err.code == "#VALUE!":
            return None
        raise


def xl_add(left: object, right: object) -> float:
    """Excel `+` with numeric coercion of both operands."""
    return _as_number(left) + _as_number(right)


def xl_sub(left: object, right: object) -> float:
    """Excel `-` with numeric coercion of both operands."""
    return _as_number(left) - _as_number(right)


def xl_mul(left: object, right: object) -> float:
    """Excel `*` with numeric coercion of both operands."""
    return _as_number(left) * _as_number(right)


def xl_div(numerator: object, denominator: object) -> float:
    """Excel `/` with `#DIV/0!` on zero and `#VALUE!` on non-numeric measures."""
    left = _as_number(numerator)
    right = _as_number(denominator)
    if right == 0:
        raise XlError("#DIV/0!")
    return left / right


def xl_pow(left: object, right: object) -> float:
    """Excel `^` with numeric coercion; overflow or complex is `#NUM!`."""
    try:
        value = _as_number(left) ** _as_number(right)
    except (ValueError, OverflowError):
        raise XlError("#NUM!") from None
    if isinstance(value, complex):
        raise XlError("#NUM!")
    return float(value)


def xl_neg(value: object) -> float:
    """Excel unary `-` with numeric coercion."""
    return -_as_number(value)


def xl_pos(value: object) -> float:
    """Excel unary `+` with numeric coercion."""
    return _as_number(value)


def _compare_key(value: object) -> tuple[int, float | str | bool]:
    """Return `(type_rank, key)` using Excel's number < text < logical order."""
    if isinstance(value, str) and is_error(value):
        raise XlError(value)
    if isinstance(value, bool):
        return 2, value
    number = _try_number(value)
    if number is not None:
        return 0, number
    if isinstance(value, str):
        return 1, value.casefold()
    raise XlError("#VALUE!")


def _cmp(left: object, right: object) -> int:
    """Return `-1`, `0`, or `1` using Excel type ranking."""
    left_rank, left_key = _compare_key(left)
    right_rank, right_key = _compare_key(right)
    if left_rank != right_rank:
        return -1 if left_rank < right_rank else 1
    if left_key == right_key:
        return 0
    if left_rank == 0:
        return -1 if float(left_key) < float(right_key) else 1
    if left_rank == 1:
        return -1 if str(left_key) < str(right_key) else 1
    return -1 if bool(left_key) < bool(right_key) else 1


def xl_eq(left: object, right: object) -> bool:
    """Excel `=`: numeric coercion when both sides are numbers, else text."""
    left_n = _try_number(left)
    right_n = _try_number(right)
    if left_n is not None and right_n is not None:
        return left_n == right_n
    if isinstance(left, str) and isinstance(right, str):
        return left.casefold() == right.casefold()
    return False


def xl_ne(left: object, right: object) -> bool:
    """Excel `<>`."""
    return not xl_eq(left, right)


def xl_lt(left: object, right: object) -> bool:
    """Excel `<` using number < text < logical ordering."""
    return _cmp(left, right) < 0


def xl_gt(left: object, right: object) -> bool:
    """Excel `>` using number < text < logical ordering."""
    return _cmp(left, right) > 0


def xl_le(left: object, right: object) -> bool:
    """Excel `<=` using number < text < logical ordering."""
    return _cmp(left, right) <= 0


def xl_ge(left: object, right: object) -> bool:
    """Excel `>=` using number < text < logical ordering."""
    return _cmp(left, right) >= 0


def xl_exp(*args: object) -> float:
    """Excel `EXP`: `e ** x`. Overflow is `#NUM!`; bad input is `#VALUE!`."""
    if len(args) != 1:
        raise XlError("#VALUE!")
    try:
        return float(math.exp(_as_number(args[0])))
    except OverflowError:
        raise XlError("#NUM!") from None


def xl_choose(index: object, *choices: float) -> float:
    """Excel `CHOOSE`: 1-based selection over already-evaluated arguments."""
    position = int(_as_number(index))
    if position < 1 or position > len(choices):
        raise XlError("#VALUE!")
    return choices[position - 1]


def xl_match(lookup: object, lookup_array: Sequence[object], match_type: int = 0) -> int:
    """Excel `MATCH` exact match (`match_type=0`), returning a 1-based index."""
    if match_type != 0:
        raise XlError("#VALUE!")
    for offset, item in enumerate(lookup_array):
        if item == lookup:
            return offset + 1
    raise XlError("#N/A")


def xl_at(values: Sequence[T], index: object) -> T:
    """Return `values[index]` (0-based), raising `#VALUE!` when out of range.

    `index` is coerced with `_as_number` and truncated toward zero, matching
    Excel INDEX/OFFSET when an arithmetic expression yields a float.
    """
    position = int(_as_number(index))
    if position < 0 or position >= len(values):
        raise XlError("#VALUE!")
    return values[position]


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
