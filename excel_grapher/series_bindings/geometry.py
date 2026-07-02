"""Row and column geometry specs for series-binding label binds.

Specs appear in bind-level `skip`/`include` lists, series-level
`exclude_rows`, and `value_map` values. A row spec is a 1-based integer or an
inclusive `"first:last"` string; a column spec is a column letter or an
inclusive `"C:D"` string.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from typing import Any, Literal

from fastpyxl.utils import column_index_from_string

__all__ = [
    "expand_column_specs",
    "expand_row_specs",
    "parse_value_map",
]

_ROW_SPEC_RE = re.compile(r"^([1-9][0-9]*)(?::([1-9][0-9]*))?$")
_COLUMN_SPEC_RE = re.compile(r"^([A-Z]{1,3})(?::([A-Z]{1,3}))?$")

Axis = Literal["rows", "columns"]


def _expand_row_spec(spec: Any) -> set[int]:
    if isinstance(spec, bool):
        raise ValueError(f"Invalid row spec: {spec!r}")
    if isinstance(spec, int):
        if spec < 1:
            raise ValueError(f"Invalid row spec: {spec!r} (rows are 1-based)")
        return {spec}
    if isinstance(spec, str):
        match = _ROW_SPEC_RE.match(spec.strip())
        if match is None:
            raise ValueError(f"Invalid row spec: {spec!r} (expected N or N:M)")
        first = int(match.group(1))
        last = int(match.group(2)) if match.group(2) else first
        if last < first:
            raise ValueError(f"Invalid row spec: {spec!r} (last row before first)")
        return set(range(first, last + 1))
    raise ValueError(f"Invalid row spec: {spec!r}")


def _expand_column_spec(spec: Any) -> set[int]:
    if not isinstance(spec, str):
        raise ValueError(f"Invalid column spec: {spec!r}")
    match = _COLUMN_SPEC_RE.match(spec.strip())
    if match is None:
        raise ValueError(f"Invalid column spec: {spec!r} (expected C or C:D)")
    first = column_index_from_string(match.group(1))
    last = column_index_from_string(match.group(2)) if match.group(2) else first
    if last < first:
        raise ValueError(f"Invalid column spec: {spec!r} (last column before first)")
    return set(range(first, last + 1))


def expand_row_specs(specs: Iterable[Any]) -> set[int]:
    """Expand row specs (ints and `"N:M"` strings) into a set of row indices."""
    rows: set[int] = set()
    for spec in specs:
        rows |= _expand_row_spec(spec)
    return rows


def expand_column_specs(specs: Iterable[Any]) -> set[int]:
    """Expand column specs (`"C"` / `"C:D"` strings) into 1-based column indices."""
    columns: set[int] = set()
    for spec in specs:
        columns |= _expand_column_spec(spec)
    return columns


def _classify_spec(spec: Any) -> Axis:
    if isinstance(spec, int) and not isinstance(spec, bool):
        return "rows"
    if isinstance(spec, str):
        if _ROW_SPEC_RE.match(spec.strip()):
            return "rows"
        if _COLUMN_SPEC_RE.match(spec.strip()):
            return "columns"
    raise ValueError(f"Invalid value_map spec: {spec!r} (expected row or column spec)")


def parse_value_map(values: Mapping[Any, Any]) -> tuple[Axis, dict[Any, set[int]]]:
    """Parse a `value_map` bind's `values` mapping into per-value index sets.

    The axis is inferred from the spec syntax: row specs (`3`, `"3:4"`) key
    data rows, column specs (`"C"`, `"C:D"`) key data columns. Mixing axes or
    assigning one row/column to two values is an error.

    Returns:
        Tuple of the inferred axis and a mapping from each manifest value to
        the set of 1-based row or column indices it covers.

    Raises:
        ValueError: On empty maps, malformed specs, mixed axes, or overlaps.
    """
    if not values:
        raise ValueError("value_map requires a non-empty values mapping")

    axis: Axis | None = None
    parsed: dict[Any, set[int]] = {}
    claimed: dict[int, Any] = {}
    for value, spec in values.items():
        specs = spec if isinstance(spec, list) else [spec]
        indices: set[int] = set()
        for item in specs:
            item_axis = _classify_spec(item)
            if axis is None:
                axis = item_axis
            elif item_axis != axis:
                raise ValueError(
                    f"value_map mixes row and column specs (value {value!r}: {item!r})"
                )
            indices |= _expand_row_spec(item) if item_axis == "rows" else _expand_column_spec(item)
        unit = "row" if axis == "rows" else "column"
        for index in indices:
            if index in claimed:
                raise ValueError(
                    f"value_map assigns {unit} {index} to both {claimed[index]!r} and {value!r}"
                )
            claimed[index] = value
        parsed[value] = indices

    if axis is None:
        raise ValueError("value_map requires at least one row or column spec")
    return axis, parsed
