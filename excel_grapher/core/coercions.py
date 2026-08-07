"""Excel-style scalar coercions and value helpers (representation-agnostic)."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from datetime import date, datetime
from typing import TypeVar, cast

# Imported for isinstance checks; stripped when coercions are embedded (Range is
# already inlined from core.grid.ranges ahead of this module).
from excel_grapher.core.grid.grid import _as_nested_rows_from_ndarray
from excel_grapher.core.grid.ranges import Range

from .types import CellValue, ExcelRange, FormulaValue, XlError

_EXCEL_EPOCH = datetime(1899, 12, 30)
T = TypeVar("T")


def _is_ndarray_like(value: object) -> bool:
    """Duck-type NumPy ndarrays without importing NumPy.

    Fast-path materialization buffers expose ``ndim`` / ``flat`` / ``tolist``.
    Matches `Grid.wrap` / `_as_nested_rows_from_ndarray` so coercions stay
    import-light for NumPy-free installs and exports.
    """
    if _as_nested_rows_from_ndarray(value) is not None:
        return True
    ndim = getattr(value, "ndim", None)
    return isinstance(ndim, int) and ndim >= 1 and hasattr(value, "flat")


_PLAIN_SCALAR_TYPES = frozenset({bool, int, float, str, type(None)})


def as_scalar(value: object) -> float | int | str | bool | XlError | None:
    """Collapse range/array values to `#VALUE!` for scalar coercion contexts.

    Lazy `Range`, unbound `ExcelRange`, and nested lists are not valid scalar
    operands. Materialized ndarray buffers (fast-path internals) also collapse
    to `#VALUE!`. Does not evaluate cells inside a `Range`.

    Cells of an exact plain scalar type return immediately: `_is_ndarray_like`
    probes several attributes that miss on every ordinary cell, and `to_number`
    / `to_string` call this once per cell in the per-cell loops.
    """
    if type(value) in _PLAIN_SCALAR_TYPES:
        return cast("float | int | str | bool | None", value)
    if isinstance(value, (Range, ExcelRange, list, tuple)) or _is_ndarray_like(value):
        return XlError.VALUE
    return cast("float | int | str | bool | XlError | None", value)


def datetime_to_excel_serial(value: datetime) -> float:
    """Convert a naive datetime to an Excel day serial (1900 date system)."""
    naive = value.replace(tzinfo=None) if value.tzinfo is not None else value
    delta = naive - _EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86_400.0


def _try_parse_iso_date_serial(text: str) -> float | None:
    stripped = text.strip()
    if not stripped:
        return None
    try:
        if "T" in stripped or " " in stripped:
            parsed = datetime.fromisoformat(stripped.replace("Z", "+00:00"))
            if parsed.tzinfo is not None:
                parsed = parsed.replace(tzinfo=None)
        else:
            parsed = datetime.combine(date.fromisoformat(stripped), datetime.min.time())
        return datetime_to_excel_serial(parsed)
    except ValueError:
        return None


def try_coerce_string_to_float(text: str) -> float | None:
    """Parse one Excel numeric string; empty/whitespace text fails (`None`)."""
    stripped = text.strip()
    if stripped == "":
        return None
    try:
        return float(stripped)
    except ValueError:
        return _try_parse_iso_date_serial(stripped)


def to_native(value: T) -> T:
    """Unwrap numpy scalars; otherwise return *value* unchanged."""
    item = getattr(value, "item", None)
    if callable(item):
        return cast("T", item())
    return value


def to_number(value: FormulaValue) -> float | XlError:
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        return scalar
    value = cast(CellValue, scalar)
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        number = try_coerce_string_to_float(value)
        if number is None:
            return XlError.VALUE
        return number
    return XlError.VALUE


def to_int(value: FormulaValue) -> int | XlError:
    """Coerce a CellValue to an integer using Excel-style numeric coercion.

    For functions that operate on integer indices (e.g. CHOOSE/INDEX/MATCH)
    while propagating Excel errors.
    """
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return int(n)


def _format_general_number(value: float | int) -> str:
    f = float(value)
    if f.is_integer():
        return str(int(f))
    return str(f)


def to_string(value: FormulaValue) -> str:
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        return scalar.value
    value = cast(CellValue, scalar)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return _format_general_number(float(value))
    if isinstance(value, str):
        return value
    return str(value)


def to_bool(value: FormulaValue) -> bool | XlError:
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        return scalar
    value = cast(CellValue, scalar)
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0
    if isinstance(value, str):
        s = value.strip().upper()
        if s == "":
            return False
        if s == "TRUE":
            return True
        if s == "FALSE":
            return False
        return XlError.VALUE
    return XlError.VALUE


def excel_casefold(value: str) -> str:
    return value.casefold()


def flatten(*args: object) -> Iterator[FormulaValue]:
    """Flatten nested lists, lazy `Range` values, and ndarray buffers in row-major order.

    Full-scan reductions (`SUM`, `COUNTIF`, …) and generic-function error
    prechecks (`get_error`) use this helper to walk multi-cell args. Selective
    consumers (`INDEX`, `MATCH`, lookups) skip `get_error` so they are not
    forced to evaluate sibling cells. Ndarray inputs are supported only for
    fast-path materialization buffers, not as persisted `CellValue` results.
    """
    for arg in args:
        if _is_ndarray_like(arg):
            flat = getattr(arg, "flat", None)
            if flat is not None:
                yield from (cast("FormulaValue", v) for v in flat)
            else:
                rows = _as_nested_rows_from_ndarray(arg)
                assert rows is not None
                yield from flatten(*rows)
            continue
        if isinstance(arg, Range):
            yield from arg.iter_raw()
            continue
        if isinstance(arg, (list, tuple)):
            yield from flatten(*arg)
            continue
        yield cast("CellValue", arg)


def get_error(*args: object) -> XlError | None:
    """Return the first flattened `XlError`, if any.

    Walks ndarrays, nested lists, and lazy `Range` cells via `flatten`. Lookup
    functions skip this precheck so selective Grid access stays consumer-driven.
    """
    for v in flatten(*args):
        if isinstance(v, XlError):
            return v
    return None


def numeric_values(values: Iterable[FormulaValue]) -> tuple[list[float], XlError | None]:
    nums: list[float] = []
    for v in values:
        n = to_number(v)
        if isinstance(n, XlError):
            return ([], n)
        nums.append(float(n))
    return (nums, None)
