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
    _schedule: object | None = field(default=None, repr=False, compare=False)

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


def _access_run_holds(
    meta: Sequence[tuple[str | None, tuple[tuple[str, int | None], ...]]],
    start: int,
    stop: int,
) -> bool:
    """True when `meta[start:stop]` share shape and each producer slot is affine."""
    shape0, pairs0 = meta[start]
    producers = tuple(producer_id for producer_id, _ in pairs0)
    n_slots = len(pairs0)
    none_slots = [False] * n_slots
    int_slots: list[list[tuple[int, int]]] = [[] for _ in range(n_slots)]
    for index in range(start, stop):
        shape, pairs = meta[index]
        if shape != shape0:
            return False
        if tuple(producer_id for producer_id, _ in pairs) != producers:
            return False
        if len(pairs) != n_slots:
            return False
        for slot, (_producer_id, prod_idx) in enumerate(pairs):
            if prod_idx is None:
                none_slots[slot] = True
            else:
                int_slots[slot].append((index, prod_idx))
    if stop - start == 1:
        return True
    for slot in range(n_slots):
        if none_slots[slot]:
            if int_slots[slot]:
                return False
            continue
        if fit_affine_map(int_slots[slot]) is None:
            return False
    return True


def _shape_partition(
    series: BoundSeries,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> tuple[Statement, ...]:
    meta = [
        (_formula_shape_key(graph, address), _cell_access_pairs(catalog, graph, address))
        for address in series.cells
    ]
    if not meta:
        return (_whole_statement(series.series_id, series.cells, series.domain),)
    runs: list[tuple[str | None, int, int]] = []
    run_start = 0
    for index in range(1, len(meta)):
        if _access_run_holds(meta, run_start, index + 1):
            continue
        runs.append((meta[run_start][0], run_start, index))
        run_start = index
    runs.append((meta[run_start][0], run_start, len(meta)))
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
        series_id: replace(series, statements=_shape_partition(series, catalog, graph))
        for series_id, series in catalog.series.items()
    }
    return SeriesCatalog(
        series=series_map,
        order=catalog.order,
        address_to_id=catalog.address_to_id,
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
