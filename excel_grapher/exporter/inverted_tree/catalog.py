"""Bound-series catalog for inverted-tree codegen."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal, cast

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    UnaryOpNode,
    resolve_cell_ref,
)
from excel_grapher.core.formula_shape import fingerprint_formula_shape
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
from excel_grapher.series_bindings.resolve import resolve_key_domain
from excel_grapher.series_bindings.types import Scalar, WorkbookSeriesBindings

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

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
    cells: tuple[str, ...]
    domain: tuple[KeyPoint, ...]


@dataclass(frozen=True, slots=True)
class BoundSeries:
    """One bindings-catalog series with expanded cells and an index domain."""

    series_id: str
    layout: Layout
    direction: Direction
    cells: tuple[str, ...]
    key_fields: tuple[str, ...]
    dtype: str
    compute_name: str | None
    raw: Mapping[str, Any]
    domain: tuple[KeyPoint, ...]
    statements: tuple[Statement, ...]
    _cell_indices: dict[str, int] = field(init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "_cell_indices",
            {normalize_address(cell): idx for idx, cell in enumerate(self.cells)},
        )

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
        return self._cell_indices.get(normalize_address(address))


def _join_key(point: KeyPoint, fields: Sequence[str]) -> tuple[Scalar, ...] | None:
    """Return `point` projected onto `fields`, or None if a field is missing."""
    try:
        return tuple(point[field] for field in fields)
    except KeyError:
        return None


def _compute_preferred_fields(series: BoundSeries) -> tuple[str, ...] | None:
    """Return join fields for `series` when every member's KeyPoint resolves.

    The preferred fields are the full key tuple, so a matrix is a loop nest —
    outer over the leading key fields, inner over `TIME_PERIOD` — and the
    identity join is unambiguous by construction (#612). `TIME_PERIOD` alone
    is the fallback when a non-time key field does not resolve.
    """
    if not series.key_fields or len(series.domain) != len(series.cells):
        return None
    if all(_join_key(point, series.key_fields) is not None for point in series.domain):
        return tuple(series.key_fields)
    if "TIME_PERIOD" in series.key_fields and all(
        _join_key(point, ("TIME_PERIOD",)) is not None for point in series.domain
    ):
        return ("TIME_PERIOD",)
    return None


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
    so identity joins do not rescan producer cells.
    """

    preferred: dict[str, tuple[str, ...] | None]
    coord_of: dict[str, int]
    index_by_coord: dict[str, dict[int, tuple[int, ...]]]


def build_schedule_index(series: Mapping[str, BoundSeries]) -> ScheduleIndex:
    """Walk bound series once and index join keys and per-cell coordinates."""
    preferred = {series_id: _compute_preferred_fields(item) for series_id, item in series.items()}
    positions: dict[tuple[str, ...], dict[tuple[Scalar, ...], int]] = {}
    for fields in {item for item in preferred.values() if item is not None}:
        ordered = _ordered_domain(series, fields)
        if ordered is None:
            continue
        positions[fields] = {key: index for index, key in enumerate(ordered)}
    coord_of: dict[str, int] = {}
    for item in series.values():
        fields = preferred[item.series_id]
        if fields is None:
            for index, address in enumerate(item.cells):
                coord_of[normalize_address(address)] = index
            continue
        lookup = positions.get(fields)
        for index, address in enumerate(item.cells):
            coord = index
            if lookup is not None and index < len(item.domain):
                key = _join_key(item.domain[index], fields)
                if key is not None and key in lookup:
                    coord = lookup[key]
            coord_of[normalize_address(address)] = coord
    index_by_coord: dict[str, dict[int, tuple[int, ...]]] = {}
    for item in series.values():
        buckets: dict[int, list[int]] = {}
        for member_index, address in enumerate(item.cells):
            coord = coord_of[normalize_address(address)]
            buckets.setdefault(coord, []).append(member_index)
        index_by_coord[item.series_id] = {
            coord: tuple(members) for coord, members in buckets.items()
        }
    return ScheduleIndex(preferred=preferred, coord_of=coord_of, index_by_coord=index_by_coord)


def preferred_fields(series: BoundSeries, catalog: SeriesCatalog) -> tuple[str, ...] | None:
    """Return cached join fields for `series` from `catalog.schedule`."""
    return catalog.schedule.preferred.get(series.series_id)


def schedule_coord(address: str, catalog: SeriesCatalog) -> int:
    """Return the member's position in the catalog's ordered index domain.

    When the cell's `KeyPoint` resolves, the coordinate is the position of the
    full key tuple in the joined domain — a loop nest with `TIME_PERIOD` as
    the inner schedule axis (#612). When only `TIME_PERIOD` resolves, its
    position is the coordinate. Otherwise the coordinate is the member's
    expansion-order index.

    Raises:
        InvertedTreeExportError: `address` is not a bound catalog cell.
    """
    found = catalog.schedule.coord_of.get(normalize_address(address))
    if found is None:
        raise InvertedTreeExportError(f"cell {address} has no schedule coordinate")
    return found


@dataclass(frozen=True, slots=True)
class SeriesCatalog:
    """Bindings series keyed by id, with reverse address lookup.

    `schedule` is built once in `build_catalog` (and copied by
    `partition_catalog`). Join coordinates are a catalog property, not a
    lazily cached attribute of scheduling.
    """

    series: dict[str, BoundSeries]
    order: tuple[str, ...]
    address_to_id: dict[str, str]
    schedule: ScheduleIndex = field(repr=False, compare=False)

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


def _key_point(values: Mapping[str, Scalar], key_fields: tuple[str, ...]) -> KeyPoint:
    return KeyPoint(tuple((field, values[field]) for field in key_fields if field in values))


def _whole_statement(
    series_id: str,
    cells: tuple[str, ...],
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


def _formula_shape_key(graph: DependencyGraph, address: str) -> str | None:
    node = graph.get_node(address)
    ast = getattr(node, "formula_ast", None) if node is not None else None
    if ast is None:
        return None
    return fingerprint_formula_shape(ast).shape_key


def _iter_cell_refs(node: object) -> Iterable[CellRefNode]:
    match node:
        case CellRefNode():
            yield node
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
    address: str,
) -> tuple[tuple[str, int | None], ...]:
    """Return `(producer_id, catalog_index)` for each cell ref at `address`."""
    node = graph.get_node(normalize_address(address))
    ast = getattr(node, "formula_ast", None) if node is not None else None
    if ast is None:
        return ()
    found: list[tuple[str, int | None]] = []
    for ref in _iter_cell_refs(ast):
        resolved = resolve_cell_ref(ref, address)
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
            shape == active_shape
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
    series_map = {
        series_id: (
            replace(series, statements=_shape_partition(series, catalog, graph))
            if series.is_formula_series
            else series
        )
        for series_id, series in catalog.series.items()
    }
    return SeriesCatalog(
        series=series_map,
        order=catalog.order,
        address_to_id=catalog.address_to_id,
        schedule=catalog.schedule,
    )


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
    partition into one `Statement` per consecutive run.
    """
    series_map: dict[str, BoundSeries] = {}
    order: list[str] = []
    address_to_id: dict[str, str] = {}
    concept_scheme = bindings.get("concept_scheme")
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
        cell_tuple = tuple(cells)
        key_fields = _key_fields_of(entry)
        try:
            raw_domain = resolve_key_domain(
                workbook,
                entry,
                cell_tuple,
                concept_scheme=concept_scheme,
                graph=graph,
            )
        except ValueError as exc:
            raise InvertedTreeExportError(str(exc)) from exc
        domain = tuple(_key_point(values, key_fields) for values in raw_domain)
        bound = BoundSeries(
            series_id=series_id,
            layout=_layout_of(entry),
            direction=_direction_of(entry),
            cells=cell_tuple,
            key_fields=key_fields,
            dtype=_dtype_of(entry),
            compute_name=_compute_name_of(entry, series_id),
            raw=entry,
            domain=domain,
            statements=(_whole_statement(series_id, cell_tuple, domain),),
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
