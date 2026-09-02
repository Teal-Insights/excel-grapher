"""Bound-series catalog for inverted-tree codegen."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, cast

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings.normalize import (
    has_constant_direction,
    has_input_direction,
    has_internal_direction,
    has_output_direction,
)
from excel_grapher.series_bindings.ranges import (
    apply_series_excludes,
    expand_data_range,
    series_data_ranges,
)
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

Direction = Literal["input", "constant", "internal", "output"]
Layout = Literal["scalar", "series", "matrix"]

_DTYPE_READ = {
    "int": "int",
    "integer": "int",
    "float": "float",
    "number": "float",
    "string": "str",
    "str": "str",
    "bool": "bool",
}


@dataclass(frozen=True, slots=True)
class BoundSeries:
    """One bindings-catalog series with expanded cell addresses."""

    series_id: str
    layout: Layout
    direction: Direction
    cells: tuple[str, ...]
    key_fields: tuple[str, ...]
    dtype: str
    compute_name: str | None
    raw: Mapping[str, Any]

    @property
    def is_scalar(self) -> bool:
        """True when the series is a single value."""
        return self.layout == "scalar" or len(self.cells) == 1

    @property
    def is_sequence(self) -> bool:
        """True when callers pass this series as a `Sequence`."""
        return not self.is_scalar

    @property
    def is_time_series(self) -> bool:
        """True when the series is a 1-D `TIME_PERIOD` sequence.

        Country×year `layout: matrix` series include `TIME_PERIOD` in
        `key_fields` but are not treated as 1-D year prefixes.
        """
        return "TIME_PERIOD" in self.key_fields and self.layout == "series"

    @property
    def is_formula_series(self) -> bool:
        """True when inverted codegen emits a helper for this series."""
        return self.direction in {"internal", "output"}

    @property
    def python_dtype(self) -> str:
        """Annotation fragment for a scalar of this series (`float`, `int`, …)."""
        return _DTYPE_READ.get(self.dtype, "float")

    def index_of(self, address: str) -> int | None:
        """Return the 0-based index of `address` in `cells`, if present."""
        try:
            return self.cells.index(normalize_address(address))
        except ValueError:
            return None


@dataclass(frozen=True, slots=True)
class SeriesCatalog:
    """Bindings series keyed by id, with reverse address lookup."""

    series: dict[str, BoundSeries]
    order: tuple[str, ...]
    address_to_id: dict[str, str]

    def get(self, series_id: str) -> BoundSeries:
        """Return the series named `series_id`."""
        try:
            return self.series[series_id]
        except KeyError as exc:
            raise InvertedTreeExportError(f"unknown series {series_id!r}") from exc

    def series_id_for(self, address: str) -> str | None:
        """Return the bound series owning `address`, if any."""
        return self.address_to_id.get(normalize_address(address))

    def series_for(self, address: str) -> BoundSeries | None:
        """Return the bound series owning `address`, if any."""
        series_id = self.series_id_for(address)
        return None if series_id is None else self.series[series_id]

    def require_series_for(self, address: str) -> BoundSeries:
        """Return the series owning `address`, or fail closed."""
        found = self.series_for(address)
        if found is None:
            raise InvertedTreeExportError(f"cell {address} is not in any bound series")
        return found

    def formula_series(self) -> list[BoundSeries]:
        """Return internals and outputs in bindings order."""
        return [self.series[sid] for sid in self.order if self.series[sid].is_formula_series]

    def output_series(self) -> list[BoundSeries]:
        """Return output series in bindings order."""
        return [self.series[sid] for sid in self.order if self.series[sid].direction == "output"]

    def input_series(self) -> list[BoundSeries]:
        """Return mutable input series in bindings order."""
        return [self.series[sid] for sid in self.order if self.series[sid].direction == "input"]

    def constant_series(self) -> list[BoundSeries]:
        """Return constant series in bindings order."""
        return [self.series[sid] for sid in self.order if self.series[sid].direction == "constant"]

    def bound_addresses(self) -> frozenset[str]:
        """Every cell owned by a bound series."""
        return frozenset(self.address_to_id)


def _direction_of(entry: Mapping[str, Any]) -> Direction:
    if has_output_direction(cast(dict[str, Any], entry)):
        return "output"
    if has_internal_direction(cast(dict[str, Any], entry)):
        return "internal"
    if has_input_direction(cast(dict[str, Any], entry)):
        return "input"
    if has_constant_direction(cast(dict[str, Any], entry)):
        return "constant"
    raise InvertedTreeExportError(
        f"series {entry.get('id')!r} has no input/constant/internal/output direction"
    )


def _layout_of(entry: Mapping[str, Any]) -> Layout:
    """Return the catalog layout.

    `matrix` is a 1-D sequence in `expand_data_range` order (issue #599).
    Nested 2-D Python types are not emitted.
    """
    layout = str(entry.get("layout") or "scalar")
    if layout == "row_series":
        layout = "series"
    if layout not in {"scalar", "series", "matrix"}:
        raise InvertedTreeExportError(
            f"series {entry.get('id')!r} has unsupported layout {layout!r}"
        )
    return cast(Layout, layout)


def _dtype_of(entry: Mapping[str, Any]) -> str:
    structure = entry.get("structure") or {}
    measure = structure.get("measure") or {}
    raw = measure.get("dtype") or measure.get("bind", {}).get("read") or "float"
    return str(raw)


def _key_fields_of(entry: Mapping[str, Any]) -> tuple[str, ...]:
    keys = entry.get("key") or []
    return tuple(str(k) for k in keys)


def _compute_name_of(entry: Mapping[str, Any], series_id: str) -> str | None:
    output = entry.get("output") or {}
    compute = output.get("compute") if isinstance(output, dict) else None
    if isinstance(compute, dict) and compute.get("name"):
        return str(compute["name"])
    if has_output_direction(cast(dict[str, Any], entry)):
        return f"compute_{series_id}"
    return None


def build_catalog(
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
) -> SeriesCatalog:
    """Expand every series `data_range` into a lookup catalog.

    Applies series-level `exclude_rows` / `exclude_columns` before indexing,
    matching `resolve_series_binding` (issue #600).
    """
    series_map: dict[str, BoundSeries] = {}
    order: list[str] = []
    address_to_id: dict[str, str] = {}
    for entry in bindings.get("series", []):
        if not isinstance(entry, dict):
            continue
        series_id = str(entry.get("id") or "")
        if not series_id:
            raise InvertedTreeExportError("series entry missing id")
        cells: list[str] = []
        for data_range in series_data_ranges(entry):
            cells.extend(
                normalize_address(addr) for addr in expand_data_range(data_range, workbook=workbook)
            )
        cells = apply_series_excludes(cells, entry)
        bound = BoundSeries(
            series_id=series_id,
            layout=_layout_of(entry),
            direction=_direction_of(entry),
            cells=tuple(cells),
            key_fields=_key_fields_of(entry),
            dtype=_dtype_of(entry),
            compute_name=_compute_name_of(entry, series_id),
            raw=entry,
        )
        series_map[series_id] = bound
        order.append(series_id)
        for address in bound.cells:
            existing = address_to_id.get(address)
            if existing is not None and existing != series_id:
                raise InvertedTreeExportError(
                    f"cell {address} is bound to both {existing!r} and {series_id!r}"
                )
            address_to_id[address] = series_id
    return SeriesCatalog(
        series=series_map,
        order=tuple(order),
        address_to_id=address_to_id,
    )


def covering_series(
    catalog: SeriesCatalog,
    addresses: Iterable[str],
) -> BoundSeries | None:
    """Return the unique series that owns every address in `addresses`."""
    ids: set[str] = set()
    for address in addresses:
        series_id = catalog.series_id_for(address)
        if series_id is None:
            return None
        ids.add(series_id)
    if len(ids) != 1:
        return None
    return catalog.get(next(iter(ids)))
