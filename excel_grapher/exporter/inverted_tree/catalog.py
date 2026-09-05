"""Bound-series catalog for inverted-tree codegen."""

from __future__ import annotations

import warnings
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal, cast

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.address_keys import (
    CanonicalAddress,
    as_canonical,
    canonical_address,
    format_cell_key,
    format_key,
    parse_address,
    parse_cell_coords,
)
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    RangeNode,
    UnaryOpNode,
    resolve_cell_ref,
)
from excel_grapher.core.formula_shape import fingerprint_formula_shape
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings.graph_predicates import (
    is_graph_formula_node,
    is_graph_leaf,
)
from excel_grapher.series_bindings.normalize import (
    effective_validation,
    has_constant_direction,
    has_input_direction,
    has_internal_direction,
    has_output_direction,
)
from excel_grapher.series_bindings.ranges import (
    apply_series_excludes,
    expand_data_range,
    format_series_data_range,
    series_data_ranges,
)
from excel_grapher.series_bindings.resolve import (
    _structure_source_addresses,
    _WorkbookValues,
    resolve_key_domain,
)
from excel_grapher.series_bindings.types import Scalar, WorkbookSeriesBindings

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

Direction = Literal["input", "constant", "internal", "output"]
Layout = Literal["scalar", "series", "matrix"]
HoleKind = Literal["blank", "off_closure", "literal", "graph_leaf"]
_SCHEDULE_AXIS = "TIME_PERIOD"
_NONE_HOLE_KINDS = frozenset({"blank", "off_closure"})

_DTYPE_READ = {
    "int": "int",
    "integer": "int",
    "float": "float",
    "number": "float",
    "string": "str",
    "str": "str",
    "bool": "bool",
    "datetime": "datetime",
    "date": "datetime",
}


@dataclass(frozen=True, slots=True)
class KeyPoint:
    """Resolved key coordinates for one series member."""

    items: tuple[tuple[str, Scalar], ...]

    def __getitem__(self, field: str) -> Scalar:
        for name, value in self.items:
            if name == field:
                return value
        raise KeyError(field)

    def as_mapping(self) -> dict[str, Scalar]:
        """Return a new dict of field names to values."""
        return dict(self.items)


@dataclass(frozen=True, slots=True)
class Statement:
    """One formula shape over an ordered index domain.

    A bound series with mixed member formulas partitions into one statement
    per consecutive shape run. Uniform series are one statement.
    """

    statement_id: str
    series_id: str
    shape_key: str | None
    start: int
    stop: int
    cells: tuple[CanonicalAddress, ...]
    domain: tuple[KeyPoint, ...]


@dataclass(frozen=True, slots=True)
class SeriesHole:
    """One catalog cell that is not an on-graph formula."""

    index: int
    address: CanonicalAddress
    kind: HoleKind
    literal: object | None = None


@dataclass(frozen=True, slots=True)
class BoundSeries:
    """One bindings-catalog series with expanded cells and an index domain."""

    series_id: str
    layout: Layout
    direction: Direction
    cells: tuple[CanonicalAddress, ...]
    key_fields: tuple[str, ...]
    dtype: str
    compute_name: str | None
    raw: Mapping[str, Any]
    domain: tuple[KeyPoint, ...]
    statements: tuple[Statement, ...]
    holes: tuple[SeriesHole, ...] = ()
    _cell_indices: dict[CanonicalAddress, int] = field(init=False, repr=False, compare=False)
    _rect: tuple[str, int, int, int, int] | None = field(init=False, repr=False, compare=False)
    _holes_by_index: dict[int, SeriesHole] = field(init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "_cell_indices",
            {cell: idx for idx, cell in enumerate(self.cells)},
        )
        object.__setattr__(self, "_rect", _dense_rect(self.cells))
        object.__setattr__(self, "_holes_by_index", {hole.index: hole for hole in self.holes})

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

    def index_of(self, address: CanonicalAddress) -> int | None:
        """Return the 0-based index of canonical `address` in `cells`, if present."""
        return self._cell_indices.get(address)

    @property
    def hole_indices(self) -> tuple[int, ...]:
        """0-based catalog indexes of cells that are not on-graph formulas."""
        return tuple(hole.index for hole in self.holes)

    @property
    def has_none_holes(self) -> bool:
        """True when a retained hole emits `None` (blank or off-closure formula)."""
        return any(hole.kind in _NONE_HOLE_KINDS for hole in self.holes)

    def hole_at(self, index: int) -> SeriesHole | None:
        """Return the hole recorded at catalog `index`, if any."""
        return self._holes_by_index.get(index)

    @property
    def rect(self) -> tuple[str, int, int, int, int] | None:
        """Axis-aligned dense rectangle `(sheet, row1, col1, row2, col2)`.

        `None` when `cells` is empty, spans sheets, or has holes.
        """
        return self._rect

    @property
    def block_width(self) -> int:
        """Column count of the row-major bound block.

        A 1-D column is width 1. A 1-D row and a `layout: matrix` series
        use the number of cells that share the first row. INDEX/OFFSET
        lowering uses this as the flat-index row stride.
        """
        if len(self.cells) <= 1:
            return 1
        first_row = parse_cell_coords(self.cells[0])[1]
        width = 0
        for cell in self.cells:
            if parse_cell_coords(cell)[1] != first_row:
                break
            width += 1
        return max(width, 1)


def _dense_rect(
    cells: Sequence[CanonicalAddress],
) -> tuple[str, int, int, int, int] | None:
    """Return the filled rectangle of `cells`, or `None` if it is not dense."""
    if not cells:
        return None
    sheet0, row0, col0 = parse_cell_coords(cells[0])
    row1 = row2 = row0
    col1 = col2 = col0
    for cell in cells[1:]:
        sheet, row, col = parse_cell_coords(cell)
        if sheet != sheet0:
            return None
        if row < row1:
            row1 = row
        elif row > row2:
            row2 = row
        if col < col1:
            col1 = col
        elif col > col2:
            col2 = col
    if (row2 - row1 + 1) * (col2 - col1 + 1) != len(cells):
        return None
    return sheet0, row1, col1, row2, col2


def _join_key(point: KeyPoint, fields: Sequence[str]) -> tuple[Scalar, ...] | None:
    """Return `point` projected onto `fields`, or None if a field is missing."""
    try:
        return tuple(point[field] for field in fields)
    except KeyError:
        return None


def _axis_fields(
    fields: tuple[str, ...] | None,
) -> tuple[tuple[str, ...], str | None]:
    """Return `(outer_key_fields, schedule_axis)` for a declared key.

    `TIME_PERIOD` is the inner schedule axis when present; every other key
    field is the instance partition of the loop nest (#612 / #638).
    """
    if fields is None or _SCHEDULE_AXIS not in fields:
        return (), None
    return tuple(name for name in fields if name != _SCHEDULE_AXIS), _SCHEDULE_AXIS


def _compute_preferred_fields(series: BoundSeries) -> tuple[str, ...] | None:
    """Return join fields for `series`, or None when no key is declared.

    The preferred fields are the full key tuple, so a matrix is a loop nest —
    outer over the leading key fields, inner over `TIME_PERIOD` — and the
    identity join is unambiguous by construction (#612). Expansion order is
    the schedule only for `key: []`. A declared key that does not resolve on
    every member raises rather than dropping to a partial join (#620).
    """
    if not series.key_fields:
        return None
    if len(series.domain) != len(series.cells):
        raise InvertedTreeExportError(
            f"series {series.series_id!r}: key domain length {len(series.domain)} "
            f"does not match cell count {len(series.cells)}"
        )
    unresolved = [
        series.cells[index]
        for index, point in enumerate(series.domain)
        if _join_key(point, series.key_fields) is None
    ]
    if unresolved:
        raise InvertedTreeExportError(
            f"series {series.series_id!r}: key did not fully resolve for cells "
            f"{', '.join(unresolved)}"
        )
    return tuple(series.key_fields)


def _ordered_domain(
    series: Mapping[str, BoundSeries], fields: Sequence[str]
) -> tuple[tuple[Scalar, ...], ...] | None:
    """Return the sorted union of resolved join keys, or None if unsortable."""
    keys: set[tuple[Scalar, ...]] = set()
    for item in series.values():
        for point in item.domain:
            key = _join_key(point, fields)
            if key is not None:
                keys.add(key)
    if not keys:
        return None
    try:
        return tuple(sorted(keys))
    except TypeError:
        return None


@dataclass(frozen=True, slots=True)
class ScheduleIndex:
    """Precomputed schedule coordinates and join fields for one catalog.

    `index_by_coord` maps series id to schedule coordinate to member indices
    so identity joins do not rescan producer cells. `statement_id_by_coord`
    maps series id to schedule coordinate to the covering statement id so
    fused-region planning is an O(1) lookup per union index.
    `partition_of` / `axis_of` split a multi-key domain into the outer
    instance partition and the inner `TIME_PERIOD` schedule axis (#638).
    `coords_of` is the set of schedule coordinates per series so peer tests
    do not rebuild those sets on every call.
    """

    preferred: dict[str, tuple[str, ...] | None]
    coord_of: dict[CanonicalAddress, int]
    index_by_coord: dict[str, dict[int, tuple[int, ...]]]
    statement_id_by_coord: dict[str, dict[int, str]] = field(default_factory=dict)
    partition_of: dict[CanonicalAddress, tuple[Scalar, ...]] = field(default_factory=dict)
    axis_of: dict[CanonicalAddress, int] = field(default_factory=dict)
    coords_of: dict[str, frozenset[int]] = field(default_factory=dict)


def _series_statement_id_by_coord(
    item: BoundSeries,
    coord_of: Mapping[CanonicalAddress, int],
) -> dict[int, str]:
    """Map schedule coordinates of `item` to the covering statement id."""
    mapping: dict[int, str] = {}
    for stmt in item.statements:
        for cell in stmt.cells:
            mapping.setdefault(coord_of[cell], stmt.statement_id)
    if not mapping:
        for address in item.cells:
            mapping.setdefault(coord_of[address], item.series_id)
    return mapping


def _statement_id_by_coord(
    series: Mapping[str, BoundSeries],
    coord_of: Mapping[CanonicalAddress, int],
) -> dict[str, dict[int, str]]:
    """Map each series id to schedule coordinate to covering statement id."""
    return {
        item.series_id: _series_statement_id_by_coord(item, coord_of) for item in series.values()
    }


def build_schedule_index(series: Mapping[str, BoundSeries]) -> ScheduleIndex:
    """Walk bound series once and index join keys and per-cell coordinates."""
    preferred = {series_id: _compute_preferred_fields(item) for series_id, item in series.items()}
    positions: dict[tuple[str, ...], dict[tuple[Scalar, ...], int]] = {}
    for fields in {item for item in preferred.values() if item is not None}:
        ordered = _ordered_domain(series, fields)
        if ordered is None:
            continue
        positions[fields] = {key: index for index, key in enumerate(ordered)}
    coord_of: dict[CanonicalAddress, int] = {}
    for item in series.values():
        fields = preferred[item.series_id]
        if fields is None:
            for index, address in enumerate(item.cells):
                coord_of[address] = index
            continue
        lookup = positions.get(fields)
        for index, address in enumerate(item.cells):
            coord = index
            if lookup is not None and index < len(item.domain):
                key = _join_key(item.domain[index], fields)
                if key is not None and key in lookup:
                    coord = lookup[key]
            coord_of[address] = coord
    index_by_coord: dict[str, dict[int, tuple[int, ...]]] = {}
    for item in series.values():
        buckets: dict[int, list[int]] = {}
        for member_index, address in enumerate(item.cells):
            coord = coord_of[address]
            buckets.setdefault(coord, []).append(member_index)
        index_by_coord[item.series_id] = {
            coord: tuple(members) for coord, members in buckets.items()
        }
    axis_values: set[Scalar] = set()
    for item in series.values():
        _outer, axis = _axis_fields(preferred[item.series_id])
        if axis is None:
            continue
        for point in item.domain:
            try:
                axis_values.add(point[axis])
            except KeyError:
                continue
    try:
        axis_lookup = {value: index for index, value in enumerate(sorted(axis_values))}
    except TypeError:
        axis_lookup = {}
    partition_of: dict[CanonicalAddress, tuple[Scalar, ...]] = {}
    axis_of: dict[CanonicalAddress, int] = {}
    for item in series.values():
        outer, axis = _axis_fields(preferred[item.series_id])
        for index, address in enumerate(item.cells):
            addr = address
            fallback = coord_of[addr]
            if axis is None or index >= len(item.domain):
                partition_of[addr] = ()
                axis_of[addr] = fallback
                continue
            point = item.domain[index]
            try:
                part = tuple(point[name] for name in outer)
                axis_value = point[axis]
            except KeyError:
                partition_of[addr] = ()
                axis_of[addr] = fallback
                continue
            partition_of[addr] = part
            axis_of[addr] = axis_lookup.get(axis_value, fallback)
    coords_of = {series_id: frozenset(coords) for series_id, coords in index_by_coord.items()}
    return ScheduleIndex(
        preferred=preferred,
        coord_of=coord_of,
        index_by_coord=index_by_coord,
        statement_id_by_coord=_statement_id_by_coord(series, coord_of),
        partition_of=partition_of,
        axis_of=axis_of,
        coords_of=coords_of,
    )


def preferred_fields(series: BoundSeries, catalog: SeriesCatalog) -> tuple[str, ...] | None:
    """Return cached join fields for `series` from `catalog.schedule`."""
    return catalog.schedule.preferred.get(series.series_id)


def schedule_coord(address: CanonicalAddress, catalog: SeriesCatalog) -> int:
    """Return the member's position in the catalog's ordered index domain.

    When the series declares a key, the coordinate is the position of the
    full key tuple in the joined domain — a loop nest with `TIME_PERIOD` as
    the inner schedule axis (#612). When no key is declared (`key: []`), the
    coordinate is the member's expansion-order index.

    `address` must already be canonical (`BoundSeries.cells`,
    `DependenceEdge` endpoints, or `canonical_address` at a public boundary).

    Raises:
        InvertedTreeExportError: `address` is not a bound catalog cell.
    """
    found = catalog.schedule.coord_of.get(address)
    if found is None:
        raise InvertedTreeExportError(f"cell {address} has no schedule coordinate")
    return found


def schedule_partition(address: CanonicalAddress, catalog: SeriesCatalog) -> tuple[Scalar, ...]:
    """Return the outer-key tuple for `address`, or `()` when there is no nest."""
    return catalog.schedule.partition_of.get(address, ())


def schedule_axis_coord(address: CanonicalAddress, catalog: SeriesCatalog) -> int:
    """Return the inner schedule-axis coordinate of `address`.

    For a `TIME_PERIOD` nest this is the year position, shared by every
    outer-key block. Otherwise it is `schedule_coord`.
    """
    found = catalog.schedule.axis_of.get(address)
    if found is not None:
        return found
    return schedule_coord(address, catalog)


@dataclass(frozen=True, slots=True)
class SeriesCatalog:
    """Bindings series keyed by id, with reverse address lookup.

    `schedule` is built once in `build_catalog` (and copied by
    `partition_catalog`). Join coordinates are a catalog property, not a
    lazily cached attribute of scheduling.
    """

    series: dict[str, BoundSeries]
    order: tuple[str, ...]
    address_to_id: dict[CanonicalAddress, str]
    schedule: ScheduleIndex = field(repr=False, compare=False)

    def get(self, series_id: str) -> BoundSeries:
        """Return the series named `series_id`."""
        try:
            return self.series[series_id]
        except KeyError as exc:
            raise InvertedTreeExportError(f"unknown series {series_id!r}") from exc

    def series_id_for(self, address: CanonicalAddress) -> str | None:
        """Return the bound series owning canonical `address`, if any."""
        return self.address_to_id.get(address)

    def series_for(self, address: CanonicalAddress) -> BoundSeries | None:
        """Return the bound series owning canonical `address`, if any."""
        series_id = self.series_id_for(address)
        return None if series_id is None else self.series[series_id]

    def require_series_for(self, address: CanonicalAddress) -> BoundSeries:
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


def _key_point(values: Mapping[str, Scalar], key_fields: tuple[str, ...]) -> KeyPoint:
    return KeyPoint(tuple((field, values[field]) for field in key_fields if field in values))


def _whole_statement(
    series_id: str,
    cells: tuple[CanonicalAddress, ...],
    domain: tuple[KeyPoint, ...],
    *,
    shape_key: str | None = None,
) -> Statement:
    return Statement(
        statement_id=series_id,
        series_id=series_id,
        shape_key=shape_key,
        start=0,
        stop=len(cells),
        cells=cells,
        domain=domain,
    )


def _formula_shape_key(graph: DependencyGraph, address: CanonicalAddress) -> str | None:
    node = graph.get_node(address)
    ast = getattr(node, "formula_ast", None) if node is not None else None
    if ast is None:
        return None
    return fingerprint_formula_shape(ast).shape_key


def _iter_cell_refs(node: object) -> Iterable[CellRefNode]:
    match node:
        case CellRefNode():
            yield node
        case RangeNode():
            yield CellRefNode(node.start_ref)
            yield CellRefNode(node.end_ref)
        case BinaryOpNode(left=left, right=right):
            yield from _iter_cell_refs(left)
            yield from _iter_cell_refs(right)
        case UnaryOpNode(operand=operand):
            yield from _iter_cell_refs(operand)
        case FunctionCallNode(args=args):
            for arg in args:
                yield from _iter_cell_refs(arg)
        case _:
            return


def fit_affine_map(pairs: Sequence[tuple[int, int]]) -> tuple[int, int] | None:
    """Return `(coeff, offset)` when `prod = coeff * host + offset` for every pair.

    A single host is underdetermined. Two products at one host, a non-integer
    slope, or a point off the line return `None`.
    """
    by_host: dict[int, int] = {}
    for host, prod in pairs:
        previous = by_host.get(host)
        if previous is not None and previous != prod:
            return None
        by_host[host] = prod
    if len(by_host) < 2:
        return None
    hosts = sorted(by_host)
    first, second = hosts[0], hosts[1]
    delta_host = second - first
    delta_prod = by_host[second] - by_host[first]
    if delta_host == 0 or delta_prod % delta_host != 0:
        return None
    coeff = delta_prod // delta_host
    offset = by_host[first] - coeff * first
    if any(by_host[host] != coeff * host + offset for host in hosts):
        return None
    return coeff, offset


def _cell_access_pairs(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    address: CanonicalAddress,
) -> tuple[tuple[str, int | None], ...]:
    """Return `(producer_id, catalog_index)` for each cell or range endpoint."""
    node = graph.get_node(address)
    ast = getattr(node, "formula_ast", None) if node is not None else None
    if ast is None:
        return ()
    found: list[tuple[str, int | None]] = []
    for ref in _iter_cell_refs(ast):
        resolved = as_canonical(resolve_cell_ref(ref, address))
        owner_id = catalog.series_id_for(resolved)
        if owner_id is None:
            found.append(("?", None))
            continue
        found.append((owner_id, catalog.get(owner_id).index_of(resolved)))
    return tuple(found)


SlotState = tuple[int, int] | int | None


def _shape_partition(
    series: BoundSeries,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> tuple[Statement, ...]:
    """Split a formula series into consecutive affine formula-shape statements."""
    meta = [
        (_formula_shape_key(graph, address), _cell_access_pairs(catalog, graph, address))
        for address in series.cells
    ]
    if not meta:
        return (_whole_statement(series.series_id, series.cells, series.domain),)
    runs: list[tuple[str | None, int, int]] = []
    run_start = 0
    active_shape, active_pairs = meta[0]
    active_producers = tuple(producer_id for producer_id, _ in active_pairs)
    slot_states: list[SlotState] = [prod for _, prod in active_pairs]

    for index in range(1, len(meta)):
        shape, pairs = meta[index]
        can_extend = (
            active_shape is not None
            and shape is not None
            and shape == active_shape
            and len(pairs) == len(active_pairs)
            and tuple(producer_id for producer_id, _ in pairs) == active_producers
        )

        next_slot_states: list[SlotState] = []
        if can_extend:
            for slot, (_producer_id, prod_idx) in enumerate(pairs):
                state = slot_states[slot]
                if state is None:
                    if prod_idx is not None:
                        can_extend = False
                        break
                    next_slot_states.append(None)
                elif isinstance(state, int):
                    if prod_idx is None:
                        can_extend = False
                        break
                    fit = fit_affine_map(((run_start, state), (index, prod_idx)))
                    if fit is None:
                        can_extend = False
                        break
                    next_slot_states.append(fit)
                else:
                    coeff, offset = state
                    if prod_idx is None or prod_idx != coeff * index + offset:
                        can_extend = False
                        break
                    next_slot_states.append(state)

        if can_extend:
            slot_states = next_slot_states
            continue

        runs.append((active_shape, run_start, index))
        run_start = index
        active_shape, active_pairs = meta[index]
        active_producers = tuple(producer_id for producer_id, _ in active_pairs)
        slot_states = [prod for _, prod in active_pairs]

    runs.append((active_shape, run_start, len(meta)))
    if len(runs) == 1:
        key = runs[0][0]
        return (_whole_statement(series.series_id, series.cells, series.domain, shape_key=key),)
    return tuple(
        Statement(
            statement_id=f"{series.series_id}__{start}",
            series_id=series.series_id,
            shape_key=key,
            start=start,
            stop=stop,
            cells=series.cells[start:stop],
            domain=series.domain[start:stop],
        )
        for key, start, stop in runs
    )


def partition_catalog(catalog: SeriesCatalog, graph: DependencyGraph) -> SeriesCatalog:
    """Split each series into consecutive formula-shape statements."""
    series_map: dict[str, BoundSeries] = {}
    statement_id_by_coord = catalog.schedule.statement_id_by_coord
    refreshed: dict[str, dict[int, str]] | None = None
    for series_id, series in catalog.series.items():
        if not series.is_formula_series:
            series_map[series_id] = series
            continue
        partitioned = replace(series, statements=_shape_partition(series, catalog, graph))
        series_map[series_id] = partitioned
        if partitioned.statements == series.statements:
            continue
        if refreshed is None:
            refreshed = dict(statement_id_by_coord)
        refreshed[series_id] = _series_statement_id_by_coord(partitioned, catalog.schedule.coord_of)
    schedule = (
        catalog.schedule
        if refreshed is None
        else replace(catalog.schedule, statement_id_by_coord=refreshed)
    )
    return SeriesCatalog(
        series=series_map,
        order=catalog.order,
        address_to_id=catalog.address_to_id,
        schedule=schedule,
    )


def _value_is_formula(value: object) -> bool:
    """True when a non-`data_only` workbook cell stores a formula."""
    if value is None:
        return False
    if isinstance(value, str):
        return value.startswith("=")
    return getattr(value, "text", None) is not None


def _stored_formula_addresses(
    workbook: Path | str, addresses: Sequence[CanonicalAddress]
) -> frozenset[CanonicalAddress]:
    """Return addresses among `addresses` that store a workbook formula."""
    if not addresses:
        return frozenset()
    from fastpyxl import load_workbook

    wanted_by_sheet: dict[str, set[str]] = {}
    for address in addresses:
        sheet, coord = parse_address(address)
        wanted_by_sheet.setdefault(sheet, set()).add(coord)
    found: set[CanonicalAddress] = set()
    book = load_workbook(workbook, data_only=False, read_only=True)
    try:
        for sheet, wanted in wanted_by_sheet.items():
            if sheet not in book.sheetnames:
                continue
            for row in book[sheet].iter_rows():
                for cell in row:
                    coord = getattr(cell, "coordinate", None)
                    if coord is None or coord not in wanted:
                        continue
                    data_type = getattr(cell, "data_type", None)
                    if data_type == "f" or _value_is_formula(cell.value):
                        found.add(as_canonical(format_key(sheet, coord)))
    finally:
        book.close()
    return frozenset(found)


def _classify_formula_holes(
    cells: Sequence[CanonicalAddress],
    graph: DependencyGraph,
    *,
    series_id: str,
    workbook: Path | str,
    reader: _WorkbookValues,
) -> tuple[SeriesHole, ...]:
    """Record retained matrix cells that are not on-graph formulas."""
    candidates = [
        (index, cell) for index, cell in enumerate(cells) if not is_graph_formula_node(graph, cell)
    ]
    if not candidates:
        return ()
    formulas = _stored_formula_addresses(workbook, [cell for _, cell in candidates])
    holes: list[SeriesHole] = []
    for index, cell in candidates:
        if is_graph_leaf(graph, cell):
            node = graph.get_node(cell)
            value = None if node is None else node.value
            if value is not None:
                holes.append(SeriesHole(index=index, address=cell, kind="graph_leaf"))
                continue
            if cell in formulas:
                holes.append(SeriesHole(index=index, address=cell, kind="off_closure"))
                continue
            raw = reader.read(cell)
            if raw is None:
                holes.append(SeriesHole(index=index, address=cell, kind="blank"))
                continue
            raise InvertedTreeExportError(
                f"series {series_id!r} cell {cell}: graph leaf has no cached value"
            )
        if cell in formulas:
            holes.append(SeriesHole(index=index, address=cell, kind="off_closure"))
            continue
        raw = reader.read(cell)
        if raw is None:
            holes.append(SeriesHole(index=index, address=cell, kind="blank"))
        else:
            holes.append(SeriesHole(index=index, address=cell, kind="literal", literal=raw))
    return tuple(holes)


def _filter_formula_series_cells(
    cells: Sequence[CanonicalAddress],
    entry: Mapping[str, Any],
    graph: DependencyGraph,
    *,
    series_id: str,
    direction: Direction,
    layout: Layout,
) -> tuple[CanonicalAddress, ...]:
    """Keep graph formula cells for internal/output series.

    Input and constant series are left unfiltered: their `data_range` is the
    public contract. Respects the same `intersect_graph_*` flags as
    `validate_series_bindings`.

    Raises:
        InvertedTreeExportError: Filtering leaves no graph formula cells.
    """
    if direction not in {"internal", "output"}:
        return tuple(cells)
    validation = effective_validation(cast(dict[str, Any], entry))
    if direction == "internal" and not validation.get("intersect_graph_formulas", True):
        return tuple(cells)
    if direction == "output" and not validation.get("intersect_graph_nodes", True):
        return tuple(cells)
    selected = tuple(cell for cell in cells if is_graph_formula_node(graph, cell))
    skipped = len(cells) - len(selected)
    if skipped > 0 and bool(validation.get("warn_on_partial_overlap", True)):
        warnings.warn(
            f"Skipped {skipped} cell(s) in data_range not graph formula cells "
            f"(series {series_id!r})",
            UserWarning,
            stacklevel=3,
        )
    if skipped > 0 and layout == "series" and len(cells) > 1 and len(selected) == 1:
        warnings.warn(
            f"series {series_id!r}: on-graph subset is a single cell; "
            "helper return type becomes scalar",
            UserWarning,
            stacklevel=3,
        )
    if not selected:
        raise InvertedTreeExportError(
            f"series {series_id!r}: data_range has no graph formula cells"
        )
    if layout == "matrix":
        return tuple(cells)
    return selected


def build_catalog(
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path | str,
    graph: DependencyGraph | None = None,
) -> SeriesCatalog:
    """Expand every series `data_range` into a lookup catalog.

    Applies series-level `exclude_rows` / `exclude_columns` before indexing,
    matching `resolve_series_binding` (issue #600). Each series carries an
    ordered key-point domain. When `graph` is provided, mixed formula shapes
    partition into one `Statement` per consecutive run, and internal/output
    series keep only graph formula cells (issue #693). Input and constant
    series stay unfiltered. `layout: matrix` keeps hole cells so the
    rectangle, stride, and key domain stay intact (issue #696).

    Raises:
        InvertedTreeExportError: A series is missing `id`, two series share an
            id (message names both `data_range`s), two series claim the same
            cell, key-domain resolution fails, a formula series has no graph
            formula cells, or a retained matrix graph leaf has no cached value.
    """
    series_map: dict[str, BoundSeries] = {}
    order: list[str] = []
    address_to_id: dict[CanonicalAddress, str] = {}
    concept_scheme = bindings.get("concept_scheme")
    pending: list[tuple[str, dict[str, Any], tuple[CanonicalAddress, ...]]] = []
    seen_raw: dict[str, Mapping[str, Any]] = {}
    for entry in bindings.get("series", []):
        if not isinstance(entry, dict):
            continue
        series_id = str(entry.get("id") or "")
        if not series_id:
            raise InvertedTreeExportError("series entry missing id")
        if series_id in seen_raw:
            previous = seen_raw[series_id]
            raise InvertedTreeExportError(
                f"duplicate series id {series_id!r}: "
                f"{format_series_data_range(previous)} and {format_series_data_range(entry)}"
            )
        seen_raw[series_id] = entry
        cells: list[CanonicalAddress] = []
        for data_range in series_data_ranges(entry):
            cells.extend(
                canonical_address(addr) for addr in expand_data_range(data_range, workbook=workbook)
            )
        cells = [as_canonical(addr) for addr in apply_series_excludes(cells, entry)]
        if graph is not None:
            cell_tuple = _filter_formula_series_cells(
                cells,
                entry,
                graph,
                series_id=series_id,
                direction=_direction_of(entry),
                layout=_layout_of(entry),
            )
        else:
            cell_tuple = tuple(cells)
        pending.append((series_id, entry, cell_tuple))
    with _WorkbookValues(workbook) as reader:
        sources: list[str] = []
        for _series_id, entry, cell_tuple in pending:
            if entry.get("key"):
                sources.extend(_structure_source_addresses(entry, cell_tuple))
        reader.prefetch(sources, graph=graph)
        for series_id, entry, cell_tuple in pending:
            key_fields = _key_fields_of(entry)
            try:
                raw_domain = resolve_key_domain(
                    workbook,
                    entry,
                    cell_tuple,
                    concept_scheme=concept_scheme,
                    graph=graph,
                    reader=reader,
                )
            except ValueError as exc:
                raise InvertedTreeExportError(str(exc)) from exc
            domain = tuple(_key_point(values, key_fields) for values in raw_domain)
            layout = _layout_of(entry)
            direction = _direction_of(entry)
            holes: tuple[SeriesHole, ...] = ()
            if graph is not None and layout == "matrix" and direction in {"internal", "output"}:
                holes = _classify_formula_holes(
                    cell_tuple,
                    graph,
                    series_id=series_id,
                    workbook=workbook,
                    reader=reader,
                )
            bound = BoundSeries(
                series_id=series_id,
                layout=layout,
                direction=direction,
                cells=cell_tuple,
                key_fields=key_fields,
                dtype=_dtype_of(entry),
                compute_name=_compute_name_of(entry, series_id),
                raw=entry,
                domain=domain,
                statements=(_whole_statement(series_id, cell_tuple, domain),),
                holes=holes,
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
    catalog = SeriesCatalog(
        series=series_map,
        order=tuple(order),
        address_to_id=address_to_id,
        schedule=build_schedule_index(series_map),
    )
    if graph is None:
        return catalog
    return partition_catalog(catalog, graph)


def covering_series(
    catalog: SeriesCatalog,
    addresses: Iterable[CanonicalAddress],
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


def _normalize_rect(start: str, end: str) -> tuple[str, int, int, int, int]:
    """Return `(sheet, row1, col1, row2, col2)` for a same-sheet A1 range."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    return sheet1, min(row1, row2), min(col1, col2), max(row1, row2), max(col1, col2)


def _rect_contains(
    rect: tuple[str, int, int, int, int],
    sheet: str,
    row1: int,
    col1: int,
    row2: int,
    col2: int,
) -> bool:
    r_sheet, r_row1, r_col1, r_row2, r_col2 = rect
    return (
        sheet == r_sheet and r_row1 <= row1 <= row2 <= r_row2 and r_col1 <= col1 <= col2 <= r_col2
    )


def _covering_rect(
    catalog: SeriesCatalog,
    sheet: str,
    row1: int,
    col1: int,
    row2: int,
    col2: int,
    *,
    lookup: CanonicalAddress,
) -> BoundSeries | None:
    """Return the unique dense series that owns the rectangle, if any."""
    series = catalog.series_for(lookup)
    if series is None:
        return None
    rect = series.rect
    if rect is None:
        return covering_series(
            catalog,
            (
                as_canonical(format_cell_key(sheet, get_column_letter(col), row))
                for row in range(row1, row2 + 1)
                for col in range(col1, col2 + 1)
            ),
        )
    if not _rect_contains(rect, sheet, row1, col1, row2, col2):
        return None
    return series


def covering_series_of_range(
    catalog: SeriesCatalog,
    start: str,
    end: str,
) -> BoundSeries | None:
    """Return the unique series that owns every cell of `start:end`.

    Tests rectangle containment against the candidate block of the start
    corner. Does not materialize the range.
    """
    sheet, row1, col1, row2, col2 = _normalize_rect(start, end)
    return _covering_rect(catalog, sheet, row1, col1, row2, col2, lookup=as_canonical(start))


def covering_series_of_column(
    catalog: SeriesCatalog,
    start: str,
    end: str,
    col_index: int,
) -> BoundSeries | None:
    """Return the unique series that owns column `col_index` of `start:end`."""
    sheet, row1, range_col1, row2, range_col2 = _normalize_rect(start, end)
    width = range_col2 - range_col1 + 1
    if col_index < 1 or col_index > width:
        raise InvertedTreeExportError(f"INDEX column {col_index} is outside range {start}:{end}")
    col = range_col1 + col_index - 1
    lookup = as_canonical(format_cell_key(sheet, get_column_letter(col), row1))
    return _covering_rect(catalog, sheet, row1, col, row2, col, lookup=lookup)


def range_column_origin(start: str, end: str, col_index: int) -> CanonicalAddress:
    """Return the top cell of 1-based column `col_index` in `start:end`."""
    sheet, row1, range_col1, _row2, range_col2 = _normalize_rect(start, end)
    width = range_col2 - range_col1 + 1
    if col_index < 1 or col_index > width:
        raise InvertedTreeExportError(f"INDEX column {col_index} is outside range {start}:{end}")
    col = range_col1 + col_index - 1
    return as_canonical(format_cell_key(sheet, get_column_letter(col), row1))
