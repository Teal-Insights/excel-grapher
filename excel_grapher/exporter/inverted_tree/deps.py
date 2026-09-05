"""First-level series dependencies, leaf closures, and fail-closed checks."""

from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator, Sequence
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Any, Literal, cast

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.address_keys import (
    CanonicalAddress,
    as_canonical,
    canonical_address,
    format_cell_key,
    format_key,
    parse_cell_coords,
)
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AbsoluteAxis,
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    resolve_cell_ref,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
)
from excel_grapher.core.range_shorthand import expand_whole_column_deps, expand_whole_row_deps
from excel_grapher.exporter.inverted_tree.access import (
    indirect_argument_addresses,
    indirect_target_addresses,
    is_seed_access,
    unique_seed_or_none,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    covering_series,
    covering_series_of_column,
    covering_series_of_range,
    fit_affine_map,
    preferred_fields,
    range_column_origin,
    schedule_axis_coord,
    schedule_coord,
    schedule_partition,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings.normalize import is_override_input

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

AccessClass = Literal[
    "identity", "shift", "affine", "gather", "whole", "dynamic", "cross_partition"
]


def _layout_distance(
    consumer: CanonicalAddress, producer: CanonicalAddress, catalog: SeriesCatalog
) -> int:
    """Return schedule-axis distance (`axis(consumer) - axis(producer)`).

    Distance is the `TIME_PERIOD` component when both cells sit on a key
    nest (#638). Flattened full-tuple coordinates are not used here — a
    Kenya 2021 read of Kenya 2020 is distance 1, not the gap across France.
    """
    consumer_series = catalog.series_for(consumer)
    producer_series = catalog.series_for(producer)
    if (
        consumer_series is not None
        and producer_series is not None
        and preferred_fields(consumer_series, catalog) != preferred_fields(producer_series, catalog)
    ):
        consumer_index = consumer_series.index_of(consumer)
        producer_index = producer_series.index_of(producer)
        return (0 if consumer_index is None else consumer_index) - (
            0 if producer_index is None else producer_index
        )
    return schedule_axis_coord(consumer, catalog) - schedule_axis_coord(producer, catalog)


def identity_join_indices(
    host: BoundSeries,
    producer: BoundSeries,
    catalog: SeriesCatalog,
) -> tuple[int, ...]:
    """Return producer catalog slots for each host member on the schedule axis.

    A hole is `-1` when no producer member shares the host's schedule
    coordinate. The join uses the same distance as `DependenceEdge` (`0`
    means identity). The producer's coordinate -> index map is cached on
    the catalog `ScheduleIndex`.

    Raises:
        InvertedTreeExportError: Two producer members share one host
            coordinate, or a host cell has no schedule coordinate.
    """
    if preferred_fields(host, catalog) != preferred_fields(producer, catalog):
        return tuple(slot if slot < len(producer.cells) else -1 for slot in range(len(host.cells)))
    producer_by_coord = catalog.schedule.index_by_coord[producer.series_id]
    slots: list[int] = []
    for host_cell in host.cells:
        host_coord = catalog.schedule.coord_of.get(host_cell)
        if host_coord is None:
            raise InvertedTreeExportError(f"cell {host_cell} has no schedule coordinate")
        matches = producer_by_coord.get(host_coord, ())
        if len(matches) > 1:
            raise InvertedTreeExportError(
                f"identity join of {host.series_id!r} onto {producer.series_id!r} "
                f"is ambiguous at {host_cell}: duplicate schedule keys "
                f"{tuple(producer.cells[slot] for slot in matches)}"
            )
        slots.append(-1 if not matches else matches[0])
    return tuple(slots)


def _member_access(
    consumer_cell: CanonicalAddress,
    producer: BoundSeries,
    producer_cell: CanonicalAddress,
    catalog: SeriesCatalog,
) -> AccessClass:
    """Classify a cell-ref by schedule distance (`0` is identity).

    Instance labels are identity or shift. `refine_access_classes` upgrades a
    consumer-producer bundle to `affine` when catalog indices lie on
    `f(i) = a*i + b` with `a != 1`.
    """
    if producer.index_of(producer_cell) is None:
        return "gather"
    if schedule_partition(consumer_cell, catalog) != schedule_partition(producer_cell, catalog):
        return "cross_partition"
    if _layout_distance(consumer_cell, producer_cell, catalog) == 0:
        return "identity"
    return "shift"


def iter_range_addresses(start: str, end: str) -> list[CanonicalAddress]:
    """Expand a same-sheet A1 range into canonical cell addresses (row-major)."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    return [
        as_canonical(format_cell_key(sheet1, get_column_letter(col), row))
        for row in range(r1, r2 + 1)
        for col in range(c1, c2 + 1)
    ]


def iter_cross_sheet_addresses(
    start: str,
    end: str,
    graph: DependencyGraph,
) -> list[CanonicalAddress]:
    """Expand a 3-D range across workbook `sheet_order` (row-major per sheet)."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    order = graph.sheet_order
    if not order:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    try:
        first = order.index(sheet1)
        last = order.index(sheet2)
    except ValueError as exc:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported") from exc
    lo, hi = (first, last) if first <= last else (last, first)
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    return [
        as_canonical(format_cell_key(sheet, get_column_letter(col), row))
        for sheet in order[lo : hi + 1]
        for row in range(r1, r2 + 1)
        for col in range(c1, c2 + 1)
    ]


def iter_ref_addresses(
    node: AstNode,
    host_cell: CanonicalAddress,
    graph: DependencyGraph | None,
) -> list[CanonicalAddress]:
    """Expand a range, whole-column, or whole-row leaf to canonical cells.

    Same-sheet rectangles use `iter_range_addresses`. Cross-sheet ranges walk
    `graph.sheet_order` between the endpoints. Whole-column / whole-row refs
    expand to the workbook used extent in `graph.sheet_bounds`.
    """
    if isinstance(node, RangeNode):
        start = resolve_cell_ref(node.start_ref, host_cell)
        end = resolve_cell_ref(node.end_ref, host_cell)
        sheet1, _row1, _col1 = parse_cell_coords(start)
        sheet2, _row2, _col2 = parse_cell_coords(end)
        if sheet1 == sheet2:
            return iter_range_addresses(start, end)
        if graph is None:
            raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
        return iter_cross_sheet_addresses(start, end, graph)
    if isinstance(node, WholeColumnNode):
        if graph is None or graph.sheet_bounds is None:
            raise InvertedTreeExportError("whole-column ref requires workbook used-range bounds")
        sheet, letter = resolve_whole_column_ref(node, host_cell)
        return [
            as_canonical(format_key(dep_sheet, a1))
            for dep_sheet, a1 in expand_whole_column_deps(sheet, letter, graph.sheet_bounds)
        ]
    if isinstance(node, WholeRowNode):
        if graph is None or graph.sheet_bounds is None:
            raise InvertedTreeExportError("whole-row ref requires workbook used-range bounds")
        sheet, row = resolve_whole_row_ref(node, host_cell)
        return [
            as_canonical(format_key(dep_sheet, a1))
            for dep_sheet, a1 in expand_whole_row_deps(sheet, row, graph.sheet_bounds)
        ]
    raise InvertedTreeExportError(
        f"expected a range, whole-column, or whole-row ref, got {type(node).__name__}"
    )


def range_column_addresses(start: str, end: str, col_index: int) -> list[CanonicalAddress]:
    """Return the 1-based column `col_index` of the rectangle `start:end`."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    width = c2 - c1 + 1
    if col_index < 1 or col_index > width:
        raise InvertedTreeExportError(f"INDEX column {col_index} is outside range {start}:{end}")
    col = c1 + col_index - 1
    return [
        as_canonical(format_cell_key(sheet1, get_column_letter(col), row))
        for row in range(r1, r2 + 1)
    ]


def node_formula_ast(graph: DependencyGraph, address: CanonicalAddress) -> AstNode:
    """Return the formula AST for `address`, or fail closed."""
    node = graph.get_node(address)
    if node is None:
        raise InvertedTreeExportError(f"graph is missing bound cell {address}")
    ast = getattr(node, "formula_ast", None)
    if ast is None:
        raise InvertedTreeExportError(
            f"bound cell {address} has no formula AST (cannot verify first-level refs)"
        )
    return ast


def _iter_cell_ref_nodes(node: AstNode) -> Iterator[CellRefNode]:
    """Yield every `CellRefNode` in `node` (not range endpoints)."""
    match node:
        case CellRefNode():
            yield node
        case BinaryOpNode(left=left, right=right):
            yield from _iter_cell_ref_nodes(left)
            yield from _iter_cell_ref_nodes(right)
        case UnaryOpNode(operand=operand):
            yield from _iter_cell_ref_nodes(operand)
        case FunctionCallNode(args=args):
            for arg in args:
                yield from _iter_cell_ref_nodes(arg)
        case _:
            return


def _ref_shifts_with_host(ref: CellRefNode, host_cell: str) -> bool:
    """True when `ref` moves with `host_cell` on every axis that differs.

    An absolute axis (`$A$2`, `A$2`, `$A2` on the axis that changes) is a
    scalar read every period and cannot be a lag.
    """
    address = resolve_cell_ref(ref, host_cell)
    _host_sheet, host_row, host_col = parse_cell_coords(host_cell)
    _addr_sheet, addr_row, addr_col = parse_cell_coords(address)
    cell = ref.ref
    col_locked = addr_col != host_col and isinstance(cell.col, AbsoluteAxis)
    row_locked = addr_row != host_row and isinstance(cell.row, AbsoluteAxis)
    return not col_locked and not row_locked


def _has_schedule_axis(series: BoundSeries, catalog: SeriesCatalog) -> bool:
    """True when `series` declares `TIME_PERIOD` as a key field."""
    fields = preferred_fields(series, catalog)
    return fields is not None and "TIME_PERIOD" in fields


def _boundary_index(series: BoundSeries, catalog: SeriesCatalog, *, delta: int) -> int:
    """Return the catalog index of the min (`delta < 0`) or max schedule coord."""
    key = min if delta < 0 else max
    return key(range(len(series.cells)), key=lambda i: schedule_coord(series.cells[i], catalog))


def _non_peer_seed_ref(
    ast: AstNode,
    host_cell: CanonicalAddress,
    series: BoundSeries,
    catalog: SeriesCatalog,
) -> CanonicalAddress | None:
    """Return the unique relative year-0 read of a series outside the host key.

    A unique relative read of a differently keyed series is a seed only when
    it is not the same `TIME_PERIOD` as `host_cell`. Copying one matrix row
    (`B4=B2`, `C4=C2`) is an identity join on the inner axis, not a scan of
    the whole matrix sequence (#649). A scalar or unkeyed neighbor has no
    schedule axis and remains a year-0 seed.
    """
    host_fields = preferred_fields(series, catalog)
    found: list[CanonicalAddress] = []
    for ref in _iter_cell_ref_nodes(ast):
        if not _ref_shifts_with_host(ref, host_cell):
            continue
        address = as_canonical(resolve_cell_ref(ref, host_cell))
        owner = catalog.series_for(address)
        if owner is None or owner.series_id == series.series_id:
            continue
        if preferred_fields(owner, catalog) == host_fields:
            continue
        if (
            _has_schedule_axis(series, catalog)
            and _has_schedule_axis(owner, catalog)
            and schedule_axis_coord(address, catalog) == schedule_axis_coord(host_cell, catalog)
        ):
            continue
        found.append(address)
    if len(found) == 1:
        return found[0]
    return None


def _adjacent_schedule_ref(
    series: BoundSeries,
    index: int,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    delta: int,
) -> CanonicalAddress | None:
    """Return a relative other-series ref at `schedule_coord` + `delta`.

    When the host and producer do not share join fields, a unique relative
    read of that producer at the series' schedule boundary is the seed or
    terminal. An overlapping keyed peer is a lagged read, not a seed. A
    same-`TIME_PERIOD` read of a richer key nest (a matrix row) is an
    identity join, not a year-0 seed (#649).
    """
    if graph is None or index < 0 or index >= len(series.cells):
        return None
    host_cell = series.cells[index]
    try:
        ast = node_formula_ast(graph, host_cell)
    except InvertedTreeExportError:
        return None
    matched: list[CanonicalAddress] = []
    for ref in _iter_cell_ref_nodes(ast):
        if not _ref_shifts_with_host(ref, host_cell):
            continue
        address = as_canonical(resolve_cell_ref(ref, host_cell))
        owner = catalog.series_for(address)
        if owner is None or owner.series_id == series.series_id:
            continue
        if not is_seed_access(series, owner, address, host_cell, catalog, delta=delta):
            continue
        matched.append(address)
    unique = unique_seed_or_none(series, host_cell, catalog, matched, delta=delta)
    if unique is not None or matched:
        return unique
    if index != _boundary_index(series, catalog, delta=delta):
        return None
    return _non_peer_seed_ref(ast, host_cell, series, catalog)


def predecessor_address(
    series: BoundSeries,
    index: int,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
) -> CanonicalAddress | None:
    """Return the lagged predecessor of `series.cells[index]`, if any.

    For index > 0 this is the previous cell in the series. For index 0 a seed
    is a relative read of a producer with no schedule axis, or of one at
    schedule-axis coordinate host - 1. Everything else is an aligned read.
    """
    if index < 0 or index >= len(series.cells):
        return None
    if index > 0:
        return series.cells[index - 1]
    return _adjacent_schedule_ref(series, index, catalog, graph, delta=-1)


def successor_address(
    series: BoundSeries,
    index: int,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
) -> CanonicalAddress | None:
    """Return the look-ahead successor of `series.cells[index]`, if any.

    For index < len(series.cells) - 1 this is the next cell in the series. For
    the last index a terminal is a relative read of a producer with no
    schedule axis, or of one at schedule-axis coordinate host + 1.
    """
    if index < 0 or index >= len(series.cells):
        return None
    if index < len(series.cells) - 1:
        return series.cells[index + 1]
    return _adjacent_schedule_ref(series, index, catalog, graph, delta=+1)


@dataclass(frozen=True, slots=True)
class DependenceEdge:
    """One instance-level read, annotated with access class and schedule distance.

    `distance` is `axis(consumer) - axis(producer)` on the inner schedule
    axis (`TIME_PERIOD` when the key is a nest). Positive means the
    producer is an earlier period (a `pre` / lag). `access` is `identity`
    when that distance is `0` and the outer keys match, `shift` when it is
    a nonzero same-partition step, `cross_partition` when the outer keys
    differ, and `affine` when catalog-index `f(i) = coeff * i + offset`
    has `coeff != 1`. `guarded` is True when the read is guarded by a
    conditional in the dependency graph.
    """

    consumer_id: str
    producer_id: str
    consumer_cell: CanonicalAddress
    producer_cell: CanonicalAddress
    distance: int
    access: AccessClass = "identity"
    coeff: int | None = None
    offset: int | None = None
    guarded: bool = False


@dataclass(frozen=True, slots=True)
class CatalogEdges:
    """Instance-level edges for every formula series, walked once.

    `by_consumer` groups `edges` by `DependenceEdge.consumer_id` so later
    planning steps do not re-walk formula ASTs.
    """

    edges: tuple[DependenceEdge, ...]
    by_consumer: dict[str, tuple[DependenceEdge, ...]]


@dataclass
class SeriesDeps:
    """Emit-facing projection of one host's `DependenceEdge`s.

    `DependenceEdge` is the source of truth. This view groups those edges by
    access class so helpers can zip, lag-index, and scan without walking the
    edge list again. It does not retain the `DependencyGraph` it was
    projected from; `edges` is the host-local source of truth:

    - `aligned_ids` / `index_maps` / `affine_maps` — identity (or affine)
      joins already taken to the host walk
    - `lagged_ids` — a producer read at both `i` and `i-1` whose
      positions shift with the host. A constant pair (`$A$8` / `$A$9`
      from every member) is a static catalog read, not a lag (#681).
    - `lookup_ids` — `whole` / `dynamic` table reads
    - `is_scan` / `seed_id` / `scan_direction` — self-lags discharged by
      loop order. A relative other-series read at `schedule_coord` ± 1 is
      a seed; an absolute selector read by every member is a scalar
      parameter, not a scan seed. A same-`TIME_PERIOD` slice of a matrix
      is an aligned take, not a seed (#649).
    """

    host_id: str
    param_ids: tuple[str, ...]
    is_scan: bool
    seed_id: str | None
    aligned_ids: frozenset[str]
    lookup_ids: frozenset[str]
    lagged_ids: frozenset[str]
    index_maps: dict[str, tuple[int, ...]]
    affine_maps: dict[str, tuple[int, int]]
    scan_direction: Literal["forward", "reversed"] = "forward"
    edges: tuple[DependenceEdge, ...] = ()


@dataclass
class _DepCollector:
    host: BoundSeries
    catalog: SeriesCatalog
    graph: DependencyGraph | None = None
    edges: list[DependenceEdge] = field(default_factory=list)

    def _emit(
        self,
        *,
        producer_id: str,
        host_cell: CanonicalAddress,
        producer_cell: CanonicalAddress,
        access: AccessClass,
    ) -> None:
        consumer_cell = host_cell
        guarded = (
            self.graph.is_guarded(consumer_cell, producer_cell) if self.graph is not None else False
        )
        self.edges.append(
            DependenceEdge(
                consumer_id=self.host.series_id,
                producer_id=producer_id,
                consumer_cell=consumer_cell,
                producer_cell=producer_cell,
                distance=_layout_distance(consumer_cell, producer_cell, self.catalog),
                access=access,
                guarded=guarded,
            )
        )

    def emit_cell(
        self, address: CanonicalAddress, host_cell: CanonicalAddress, host_index: int
    ) -> None:
        owner = self.catalog.require_series_for(address)
        if owner.series_id == self.host.series_id:
            if address == self.host.cells[host_index]:
                return
            self._emit(
                producer_id=owner.series_id,
                host_cell=host_cell,
                producer_cell=address,
                access=_member_access(host_cell, owner, address, self.catalog),
            )
            return
        self._emit(
            producer_id=owner.series_id,
            host_cell=host_cell,
            producer_cell=address,
            access=_member_access(host_cell, owner, address, self.catalog),
        )

    def emit_lookup(
        self,
        producer: BoundSeries,
        host_cell: CanonicalAddress,
        producer_cell: CanonicalAddress,
        access: AccessClass,
    ) -> None:
        if producer.series_id == self.host.series_id:
            return
        self._emit(
            producer_id=producer.series_id,
            host_cell=host_cell,
            producer_cell=producer_cell,
            access=access,
        )

    def _visit_range_addresses(
        self,
        addresses: list[CanonicalAddress],
        host_cell: CanonicalAddress,
    ) -> None:
        if not addresses:
            raise InvertedTreeExportError(f"series {self.host.series_id!r}: empty range")
        covered = covering_series(self.catalog, addresses)
        if covered is not None:
            self.emit_lookup(covered, host_cell, addresses[0], "whole")
            return
        missing = [addr for addr in addresses if self.catalog.series_id_for(addr) is None]
        if missing:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} range is not a bound series "
                f"(unbound cells: {missing[:8]})"
            )
        seen: set[str] = set()
        for addr in addresses:
            series_id = self.catalog.series_id_for(addr)
            if series_id is None or series_id in seen:
                continue
            seen.add(series_id)
            self.emit_lookup(self.catalog.get(series_id), host_cell, addr, "whole")

    def visit(self, node: AstNode, *, host_cell: CanonicalAddress, host_index: int) -> None:
        match node:
            case CellRefNode():
                address = as_canonical(resolve_cell_ref(node, host_cell))
                self.emit_cell(address, host_cell, host_index)
            case RangeNode() | WholeColumnNode() | WholeRowNode():
                self._visit_range_addresses(
                    iter_ref_addresses(node, host_cell, self.graph),
                    host_cell,
                )
            case FunctionCallNode():
                self._visit_function(node, host_cell=host_cell, host_index=host_index)
            case BinaryOpNode():
                self.visit(node.left, host_cell=host_cell, host_index=host_index)
                self.visit(node.right, host_cell=host_cell, host_index=host_index)
            case UnaryOpNode():
                self.visit(node.operand, host_cell=host_cell, host_index=host_index)
            case _:
                return

    def _visit_function(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        name = normalize_excel_function_name(node.name)
        if name == "OFFSET":
            self._visit_offset(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "INDEX":
            self._visit_index(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "INDIRECT":
            self._visit_indirect(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "MATCH":
            self._visit_match(node, host_cell=host_cell, host_index=host_index)
            return
        for arg in node.args:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _visit_offset(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        if not node.args:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: OFFSET with no arguments"
            )
        table = self._series_for_ref(node.args[0], host_cell)
        self.emit_lookup(table, host_cell, self._ref_anchor(node.args[0], host_cell), "dynamic")
        for arg in node.args[1:]:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _visit_indirect(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        del host_index
        if self.graph is None:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: "
                "INDIRECT has no graph to classify"
            )
        exclude = indirect_argument_addresses(node, host_cell)
        targets = indirect_target_addresses(self.graph, host_cell, exclude=tuple(exclude))
        if not targets:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: INDIRECT has no resolved edges"
            )
        covered = covering_series(self.catalog, targets)
        if covered is None:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {host_cell}: "
                "INDIRECT targets are not one bound series"
            )
        self.emit_lookup(covered, host_cell, targets[0], "dynamic")

    def _visit_index(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        if len(node.args) < 2:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: INDEX expects a range and row"
            )
        row_arg = node.args[1]
        col_arg = node.args[2] if len(node.args) > 2 else None
        if isinstance(row_arg, FunctionCallNode) and (
            normalize_excel_function_name(row_arg.name) == "MATCH"
        ):
            self._visit_match(row_arg, host_cell=host_cell, host_index=host_index)
        else:
            self.visit(row_arg, host_cell=host_cell, host_index=host_index)
        col_index = 1
        col_literal = True
        if isinstance(col_arg, NumberNode):
            col_index = int(col_arg.value)
        elif col_arg is not None:
            col_literal = False
            try:
                self.visit(col_arg, host_cell=host_cell, host_index=host_index)
            except InvertedTreeExportError as exc:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r} cell {host_cell}: "
                    f"INDEX column cannot be lowered ({exc})"
                ) from exc
        if isinstance(node.args[0], RangeNode):
            start = resolve_cell_ref(node.args[0].start_ref, host_cell)
            end = resolve_cell_ref(node.args[0].end_ref, host_cell)
            covered_full = covering_series_of_range(self.catalog, start, end)
            if covered_full is not None:
                self.emit_lookup(covered_full, host_cell, as_canonical(start), "dynamic")
            elif not col_literal:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r} cell {host_cell}: "
                    "INDEX column cannot be lowered"
                )
            else:
                covered = covering_series_of_column(self.catalog, start, end, col_index)
                if covered is None:
                    raise InvertedTreeExportError(
                        f"series {self.host.series_id!r} cell {host_cell}: "
                        f"INDEX column {col_index} of {start}:{end} is not a bound series"
                    )
                self.emit_lookup(
                    covered, host_cell, range_column_origin(start, end, col_index), "dynamic"
                )
        else:
            self.visit(node.args[0], host_cell=host_cell, host_index=host_index)

    def _visit_match(
        self,
        node: FunctionCallNode,
        *,
        host_cell: CanonicalAddress,
        host_index: int,
    ) -> None:
        if len(node.args) < 2:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: MATCH expects lookup and array"
            )
        self.visit(node.args[0], host_cell=host_cell, host_index=host_index)
        array_series = self._series_for_ref(node.args[1], host_cell)
        self.emit_lookup(
            array_series, host_cell, self._ref_anchor(node.args[1], host_cell), "dynamic"
        )
        for arg in node.args[2:]:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _series_for_ref(self, node: AstNode, host_cell: CanonicalAddress) -> BoundSeries:
        if isinstance(node, CellRefNode):
            return self.catalog.require_series_for(as_canonical(resolve_cell_ref(node, host_cell)))
        if isinstance(node, (RangeNode, WholeColumnNode, WholeRowNode)):
            addresses = iter_ref_addresses(node, host_cell, self.graph)
            covered = covering_series(self.catalog, addresses)
            if covered is None:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: reference is not a bound series"
                )
            return covered
        raise InvertedTreeExportError(
            f"series {self.host.series_id!r}: expected a cell or range reference, "
            f"got {type(node).__name__}"
        )

    def _ref_anchor(self, node: AstNode, host_cell: CanonicalAddress) -> CanonicalAddress:
        if isinstance(node, CellRefNode):
            return as_canonical(resolve_cell_ref(node, host_cell))
        if isinstance(node, RangeNode):
            return as_canonical(resolve_cell_ref(node.start_ref, host_cell))
        raise InvertedTreeExportError(
            f"series {self.host.series_id!r}: expected a cell or range reference, "
            f"got {type(node).__name__}"
        )


def refine_access_classes(
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
) -> list[DependenceEdge]:
    """Rewrite identity/shift bundles that lie on an integer affine map."""
    groups: dict[tuple[str, str], list[int]] = {}
    result = list(edges)
    for index, edge in enumerate(result):
        if edge.access in {"whole", "dynamic", "gather", "cross_partition"}:
            continue
        groups.setdefault((edge.consumer_id, edge.producer_id), []).append(index)
    for (consumer_id, producer_id), indexes in groups.items():
        consumer = catalog.get(consumer_id)
        producer = catalog.get(producer_id)
        pairs: list[tuple[int, int]] = []
        valid = True
        for index in indexes:
            edge = result[index]
            host = consumer.index_of(edge.consumer_cell)
            prod = producer.index_of(edge.producer_cell)
            if host is None or prod is None:
                valid = False
                break
            pairs.append((host, prod))
        if not valid:
            continue
        fitted = fit_affine_map(pairs)
        if fitted is None or fitted[0] == 1:
            continue
        coeff, offset = fitted
        for index in indexes:
            result[index] = replace(result[index], access="affine", coeff=coeff, offset=offset)
    return result


def collect_series_edges(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> list[DependenceEdge]:
    """Walk `series` formulas and return classified instance-level edges."""
    collector = _DepCollector(host=series, catalog=catalog, graph=graph)
    for index, address in enumerate(series.cells):
        ast = node_formula_ast(graph, address)
        collector.visit(ast, host_cell=address, host_index=index)
    return refine_access_classes(collector.edges, catalog)


def requires_demand_driven(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
    edges: Sequence[DependenceEdge] | None = None,
) -> bool:
    """True when a same-series ref cannot be discharged by a scan.

    Forward self-lags whose schedule distances are all positive (`t-k`,
    including mixed `{t-1, t-2}`) use a fused scan. Unit backward recursion
    (`value_t = value_{t+1} * k`) uses a reversed scan. Mixed directions and
    irregular self-refs fall through to demand-driven.

    `edges` should be the already-collected edges of `series` (or the
    catalog). When omitted, the formulas are walked via `graph`.
    """
    if edges is None:
        if graph is None:
            raise TypeError("requires_demand_driven requires edges or graph")
        series_edges: Sequence[DependenceEdge] = collect_series_edges(
            series, catalog=catalog, graph=graph
        )
    else:
        series_edges = edges
    self_edges = [
        edge
        for edge in series_edges
        if edge.consumer_id == series.series_id and edge.producer_id == series.series_id
    ]
    if not self_edges:
        return False
    if all(edge.distance > 0 for edge in self_edges):
        return False

    is_backward = True
    for edge in self_edges:
        host_index = series.index_of(edge.consumer_cell)
        if host_index is None:
            return True
        succ = successor_address(series, host_index, catalog, graph)
        if succ is None or edge.producer_cell != succ:
            is_backward = False
            break
    return not is_backward


def collect_catalog_edges(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> CatalogEdges:
    """Walk each formula series once and return catalog-wide classified edges."""
    by_consumer: dict[str, tuple[DependenceEdge, ...]] = {}
    collected: list[DependenceEdge] = []
    for series in catalog.formula_series():
        series_edges = tuple(collect_series_edges(series, catalog=catalog, graph=graph))
        by_consumer[series.series_id] = series_edges
        collected.extend(series_edges)
    return CatalogEdges(edges=tuple(collected), by_consumer=by_consumer)


def collect_all_dependence_edges(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> tuple[DependenceEdge, ...]:
    """Collect instance-level edges from every formula series to bound producers."""
    return collect_catalog_edges(catalog, graph).edges


def series_deps_from_edges(
    host: BoundSeries,
    edges: Sequence[DependenceEdge],
    catalog: SeriesCatalog,
    graph: DependencyGraph | None = None,
) -> SeriesDeps:
    """Derive the emit-facing `SeriesDeps` view from classified edges of `host`."""
    params: dict[str, None] = {}
    lookup_ids: set[str] = set()
    aligned_hits: dict[str, list[tuple[int, int]]] = {}
    saw_self_lag = False
    saw_forward_lag = False
    saw_backward_lag = False
    seed_id: str | None = None
    terminal_seed_id: str | None = None
    for edge in edges:
        if edge.consumer_id != host.series_id:
            continue
        if edge.producer_id == host.series_id:
            host_index = host.index_of(edge.consumer_cell)
            if host_index is None:
                continue
            if edge.distance > 0:
                saw_self_lag = True
            pred = predecessor_address(host, host_index, catalog, graph)
            if pred is not None and edge.producer_cell == pred:
                saw_forward_lag = True
            succ = successor_address(host, host_index, catalog, graph)
            if succ is not None and edge.producer_cell == succ:
                saw_backward_lag = True
            continue
        params.setdefault(edge.producer_id, None)
        if edge.access in {"whole", "dynamic"}:
            lookup_ids.add(edge.producer_id)
            continue
        host_index = host.index_of(edge.consumer_cell)
        if host_index is None:
            continue
        owner = catalog.get(edge.producer_id)
        idx = owner.index_of(edge.producer_cell)
        if idx is not None:
            aligned_hits.setdefault(edge.producer_id, []).append((host_index, idx))
        pred = _adjacent_schedule_ref(host, host_index, catalog, graph, delta=-1)
        if pred is not None and edge.producer_cell == pred:
            seed_id = edge.producer_id
        succ = _adjacent_schedule_ref(host, host_index, catalog, graph, delta=+1)
        if succ is not None and edge.producer_cell == succ:
            terminal_seed_id = edge.producer_id
    scan_direction: Literal["forward", "reversed"] = "forward"
    if saw_backward_lag and not saw_forward_lag:
        scan_direction = "reversed"
        is_scan = True
        if terminal_seed_id is not None:
            seed_id = terminal_seed_id
    elif saw_self_lag or seed_id is not None:
        scan_direction = "forward"
        is_scan = True
    else:
        is_scan = False
    remaining = [sid for sid in catalog.order if sid in params]
    if seed_id is not None and seed_id in remaining:
        remaining.remove(seed_id)
        param_ids = (seed_id, *remaining)
    else:
        param_ids = tuple(remaining)
    index_maps: dict[str, tuple[int, ...]] = {}
    affine_maps: dict[str, tuple[int, int]] = {}
    aligned: set[str] = set()
    lagged: set[str] = set()
    host_n = len(host.cells)
    identity_by_producer: dict[str, list[DependenceEdge]] = {}
    for edge in edges:
        if (
            edge.consumer_id == host.series_id
            and edge.access == "identity"
            and edge.producer_id != host.series_id
        ):
            identity_by_producer.setdefault(edge.producer_id, []).append(edge)
    for series_id, pairs in aligned_hits.items():
        if series_id in lookup_ids or series_id == seed_id:
            continue
        dep = catalog.get(series_id)
        if dep.is_scalar:
            continue
        per_host: dict[int, set[int]] = {}
        for host_i, dep_i in pairs:
            per_host.setdefault(host_i, set()).add(dep_i)
        slots = [-1] * host_n
        saw_lag = False
        origin = fit_affine_map(
            [(host_i, min(indices)) for host_i, indices in per_host.items() if indices]
        )
        static_catalog = origin is not None and origin[0] == 0
        if static_catalog and any(len(indices) > 1 for indices in per_host.values()):
            # Same catalog slots from every member (`labels[0]` / `labels[1]`),
            # or a mixed absolute + relative read of one producer (#681).
            continue
        for host_i, indices in per_host.items():
            if len(indices) > 2:
                raise InvertedTreeExportError(
                    f"series {host.series_id!r} cell {host.cells[host_i]} "
                    f"reads {series_id!r} at more than two positions {tuple(sorted(indices))}"
                )
            if len(indices) == 2:
                lo, hi = sorted(indices)
                if hi - lo != 1:
                    raise InvertedTreeExportError(
                        f"series {host.series_id!r} cell {host.cells[host_i]} "
                        f"reads {series_id!r} at two positions ({hi}, {lo})"
                    )
                saw_lag = True
                slots[host_i] = hi
            else:
                slots[host_i] = next(iter(indices))
        joined = identity_join_indices(host, dep, catalog)
        for edge in identity_by_producer.get(series_id, ()):
            host_index = host.index_of(edge.consumer_cell)
            if host_index is None:
                continue
            slot = joined[host_index]
            if slot < 0 or dep.cells[slot] != edge.producer_cell:
                raise InvertedTreeExportError(
                    f"series {host.series_id!r} cell {edge.consumer_cell} "
                    f"identity-reads {edge.producer_cell}, not the join slot "
                    f"of {series_id!r}"
                )
        if saw_lag:
            lagged.add(series_id)
            continue
        if not all(slot >= 0 for slot in slots):
            continue
        fitted = fit_affine_map([(index, slots[index]) for index in range(host_n)])
        if tuple(slots) == joined:
            index_maps[series_id] = joined
            aligned.add(series_id)
            if fitted is not None and fitted[0] != 1:
                affine_maps[series_id] = fitted
        elif fitted is not None and (
            fitted[0] != 1 or preferred_fields(host, catalog) != preferred_fields(dep, catalog)
        ):
            index_maps[series_id] = tuple(slots)
            aligned.add(series_id)
            if fitted[0] != 1:
                affine_maps[series_id] = fitted
    return SeriesDeps(
        host_id=host.series_id,
        param_ids=param_ids,
        is_scan=is_scan,
        seed_id=seed_id,
        aligned_ids=frozenset(aligned),
        lookup_ids=frozenset(lookup_ids),
        lagged_ids=frozenset(lagged),
        index_maps=index_maps,
        affine_maps=affine_maps,
        scan_direction=scan_direction,
        edges=tuple(edge for edge in edges if edge.consumer_id == host.series_id),
    )


def collect_series_deps(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> SeriesDeps:
    """Collect first-level bound-series dependencies of `series`."""
    return series_deps_from_edges(
        series,
        collect_series_edges(series, catalog=catalog, graph=graph),
        catalog,
        graph,
    )


def collect_all_deps(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    catalog_edges: CatalogEdges | None = None,
) -> dict[str, SeriesDeps]:
    """Collect first-level deps for every formula series.

    Pass `catalog_edges` when the catalog has already been walked so this
    step does not re-visit formula ASTs.
    """
    if catalog_edges is None:
        catalog_edges = collect_catalog_edges(catalog, graph)
    return {
        series.series_id: series_deps_from_edges(
            series, catalog_edges.by_consumer.get(series.series_id, ()), catalog, graph
        )
        for series in catalog.formula_series()
    }


def leaf_closure(
    root_id: str,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    """Return input and constant series in the subgraph of `root_id`.

    Order follows the bindings catalog (`catalog.order`).
    """
    seen: set[str] = set()
    stack = [root_id]
    while stack:
        current = stack.pop()
        if current in seen:
            continue
        seen.add(current)
        info = deps.get(current)
        if info is None:
            continue
        for param_id in info.param_ids:
            stack.append(param_id)
    inputs = [sid for sid in catalog.order if sid in seen and catalog.get(sid).direction == "input"]
    constants = [
        sid for sid in catalog.order if sid in seen and catalog.get(sid).direction == "constant"
    ]
    return tuple(inputs + constants)


def predecessor_closure(
    indices: Sequence[int],
    distances: Sequence[int] = (1,),
) -> tuple[int, ...]:
    """Return `indices` closed under each positive lag in `distances`.

    The default `distances=(1,)` is the unit-lag scan: wanting `{2, 4}` needs
    `{0, 1, 2, 3, 4}`. A stride-`k` recurrence passes `distances=(k,)`. The
    closure is the lag graph, not a string-processing `1..max` rule.
    """
    from excel_grapher.exporter.inverted_tree.schedule import IndexSet

    return IndexSet.from_indices(indices).closure_under(distances).materialize()


def plan_indices(
    output: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]] | None = None,
) -> tuple[dict[str, tuple[int, ...]], dict[str, tuple[int, ...]]]:
    """Plan catalog indices each series must yield, working backward from `output`.

    Returns `(result_indices, call_indices)`:

    - `result_indices[id]`: catalog positions consumers need from `id`
    - `call_indices[id]`: positions a formula series actually computes (scan
      edges expand via `predecessor_closure`)

    Lookup / constrained tables stay identity (full catalog). A scan that
    consumes an elementwise series wins: the scan's closure is unioned into
    that series' needed set.

    Multi-series lag zippers are one call: every member is computed for the
    full catalog (the year loop cannot skip a prefix). Consumers then `take`.
    A nested fused nest indexes external producers with `_area` in catalog
    space, so those leaves stay identity too.
    """
    result: dict[str, tuple[int, ...]] = {
        output.series_id: tuple(range(len(output.cells))),
    }
    call: dict[str, tuple[int, ...]] = {}

    def add_result(series_id: str, indices: tuple[int, ...]) -> None:
        previous = result.get(series_id)
        if previous is None:
            result[series_id] = tuple(sorted(set(indices)))
            return
        result[series_id] = tuple(sorted(set(previous) | set(indices)))

    formula_ids = formula_closure(output.series_id, catalog=catalog, deps=deps, scc_map=scc_map)
    processed_sccs: set[tuple[str, ...]] = set()
    for host_id in reversed(formula_ids):
        scc = (scc_map or {}).get(host_id, (host_id,))
        if len(scc) > 1:
            if scc in processed_sccs:
                continue
            processed_sccs.add(scc)
            if not any(sid in result for sid in scc):
                continue
            members = set(scc)
            for sid in scc:
                full = tuple(range(len(catalog.get(sid).cells)))
                if sid not in result:
                    add_result(sid, full)
                call[sid] = full
            for sid in scc:
                _propagate_param_indices(
                    deps[sid],
                    call[sid],
                    catalog=catalog,
                    skip=members,
                    add_result=add_result,
                )
            if _scc_is_nested(scc, catalog):
                for sid in scc:
                    for param_id in deps[sid].param_ids:
                        if param_id in members:
                            continue
                        dep = catalog.get(param_id)
                        if dep.is_sequence:
                            add_result(param_id, tuple(range(len(dep.cells))))
            continue
        info = deps[host_id]
        if host_id not in result:
            continue
        host_result = result[host_id]
        host_call = predecessor_closure(host_result) if info.is_scan else host_result
        call[host_id] = host_call
        _propagate_param_indices(
            info,
            host_call,
            catalog=catalog,
            skip=set(),
            add_result=add_result,
        )
    return result, call


def _scc_is_nested(scc: tuple[str, ...], catalog: SeriesCatalog) -> bool:
    """True when `scc` spans more than one outer-key partition."""
    seen: set[tuple[object, ...]] = set()
    for sid in scc:
        for address in catalog.get(sid).cells:
            part = schedule_partition(address, catalog)
            if part:
                seen.add(part)
            if len(seen) > 1:
                return True
    return False


def _propagate_param_indices(
    info: SeriesDeps,
    host_call: tuple[int, ...],
    *,
    catalog: SeriesCatalog,
    skip: set[str],
    add_result: Callable[[str, tuple[int, ...]], None],
) -> None:
    """Union consumer indices into each first-level param of `info`."""
    for param_id in info.param_ids:
        if param_id in skip:
            continue
        dep = catalog.get(param_id)
        if param_id in info.lookup_ids:
            add_result(param_id, tuple(range(len(dep.cells))))
            continue
        if dep.is_scalar:
            add_result(param_id, tuple(range(len(dep.cells))))
            continue
        affine = info.affine_maps.get(param_id)
        if affine is not None:
            from excel_grapher.exporter.inverted_tree.schedule import IndexSet

            coeff, offset = affine
            add_result(
                param_id,
                IndexSet.from_indices(host_call).map_affine(coeff, offset).materialize(),
            )
            continue
        index_map = info.index_maps.get(param_id)
        if index_map is None:
            add_result(param_id, tuple(range(len(dep.cells))))
            continue
        add_result(param_id, tuple(index_map[index] for index in host_call))


def formula_closure(
    root_id: str,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]] | None = None,
) -> tuple[str, ...]:
    """Return formula series in the subgraph of `root_id` (bindings order, topo).

    Period-lag zipper SCCs are a single unit so they do not fail the DAG sort.
    """
    seen: set[str] = set()
    stack = [root_id]
    while stack:
        current = stack.pop()
        if current in seen:
            continue
        seen.add(current)
        info = deps.get(current)
        if info is None:
            continue
        for param_id in info.param_ids:
            if catalog.get(param_id).is_formula_series:
                stack.append(param_id)
    formula_ids = [
        sid for sid in catalog.order if sid in seen and catalog.get(sid).is_formula_series
    ]
    if scc_map is None:
        return tuple(_topo_sort(formula_ids, deps=deps))
    units: list[tuple[str, ...]] = []
    seen_units: set[tuple[str, ...]] = set()
    for sid in formula_ids:
        unit = scc_map.get(sid, (sid,))
        if unit not in seen_units:
            seen_units.add(unit)
            units.append(unit)
    return tuple(sid for unit in _topo_units(units, deps=deps, scc_map=scc_map) for sid in unit)


def _topo_units(
    units: Sequence[tuple[str, ...]],
    *,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]],
) -> list[tuple[str, ...]]:
    """Topologically sort SCC supernodes (dependencies first)."""
    selected = set(units)
    remaining = set(units)
    ordered: list[tuple[str, ...]] = []
    while remaining:
        ready = [
            unit
            for unit in units
            if unit in remaining
            and all(
                scc_map.get(pid, (pid,)) not in remaining
                for sid in unit
                for pid in (deps[sid].param_ids if sid in deps else ())
                if scc_map.get(pid, (pid,)) in selected and scc_map.get(pid, (pid,)) != unit
            )
        ]
        if not ready:
            raise InvertedTreeExportError(
                f"cyclic formula-series dependencies among "
                f"{sorted({sid for unit in remaining for sid in unit})}"
            )
        for unit in ready:
            remaining.remove(unit)
            ordered.append(unit)
    return ordered


def _topo_sort(series_ids: Iterable[str], *, deps: dict[str, SeriesDeps]) -> list[str]:
    selected = set(series_ids)
    remaining = set(selected)
    ordered: list[str] = []
    while remaining:
        ready = [
            sid
            for sid in series_ids
            if sid in remaining
            and all(
                pid not in remaining
                for pid in (deps[sid].param_ids if sid in deps else ())
                if pid in selected
            )
        ]
        if not ready:
            raise InvertedTreeExportError(
                f"cyclic formula-series dependencies among {sorted(remaining)}"
            )
        for sid in ready:
            remaining.remove(sid)
            ordered.append(sid)
    return ordered


def assert_subgraph_bound(
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    roots: Iterable[str],
) -> None:
    """Fail closed when a subgraph formula is unbound or bound as a plain input."""
    seen: set[str] = set()
    stack = [canonical_address(addr) for addr in roots]
    while stack:
        address = stack.pop()
        if address in seen:
            continue
        seen.add(address)
        node = graph.get_node(address)
        if node is None:
            continue
        has_formula = bool(getattr(node, "has_formula", False))
        if has_formula:
            owner = catalog.series_for(address)
            if owner is None:
                raise InvertedTreeExportError(
                    f"unbound formula cell {address} is in the target subgraph"
                )
            if not owner.is_formula_series and not (
                owner.direction == "input" and is_override_input(cast(dict[str, Any], owner.raw))
            ):
                raise InvertedTreeExportError(
                    f"formula cell {address} is bound to {owner.direction} series "
                    f"{owner.series_id!r}, not an internal or output series"
                )
        for dep in graph.get_dependencies(address):
            stack.append(as_canonical(dep))


def all_formula_root_cells(catalog: SeriesCatalog) -> Iterator[str]:
    """Yield every cell of every formula series."""
    for series in catalog.formula_series():
        yield from series.cells
