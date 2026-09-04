"""Access functions derived from the graph's resolved edges.

Lookups classify one function per `(host statement, producer block)`.
Plain cell references classify one function per formula site of that
producer — a mixed relative and absolute read is two accesses. Each
producer axis is `static` (affine in the host index), `dynamic`
(candidate set + runtime selector), or `whole` (the block itself).
Anything else fails closed.
"""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    UnaryOpNode,
    resolve_cell_ref,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    fit_affine_map,
    preferred_fields,
    schedule_axis_coord,
    schedule_coord,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.dependency_provenance import DependencyCause

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

AxisKind = Literal["static", "dynamic", "whole"]


@dataclass(frozen=True, slots=True)
class AxisAccess:
    """One producer-axis classification.

    `coeff` / `offset` map the host index to the origin (minimum coordinate)
    of the candidate set on this axis. A static singleton's origin is the
    member itself. A sliding window stores the window start.
    """

    kind: AxisKind
    coeff: int = 0
    offset: int = 0

    def linear_terms(self) -> tuple[int, int]:
        """Return `(coeff, offset)` for the candidate-set origin."""
        return self.coeff, self.offset


@dataclass(frozen=True, slots=True)
class AccessFunction:
    """Row/column access of `producer_id` from `host_id`."""

    host_id: str
    producer_id: str
    row: AxisAccess
    col: AxisAccess
    width: int

    def flat_index_expr(self, row_expr: str, col_expr: str) -> str:
        """Return `row_expr * W + col_expr` for a row-major block."""
        if self.width <= 1:
            return row_expr if col_expr in {"0", "0.0"} else f"{row_expr} + {col_expr}"
        if col_expr in {"0", "0.0"}:
            return f"{row_expr} * {self.width}"
        return f"{row_expr} * {self.width} + {col_expr}"


def _canonical(address: str) -> CanonicalAddress:
    return as_canonical(address)


def indirect_argument_addresses(
    node: FunctionCallNode, host_cell: CanonicalAddress
) -> set[CanonicalAddress]:
    """Return cell addresses that feed `INDIRECT`'s text or A1-flag arguments."""
    found: set[CanonicalAddress] = set()
    for arg in node.args:
        if isinstance(arg, CellRefNode):
            found.add(as_canonical(resolve_cell_ref(arg, host_cell)))
    return found


def indirect_target_addresses(
    graph: DependencyGraph,
    host_cell: CanonicalAddress,
    *,
    exclude: Sequence[CanonicalAddress] = (),
) -> list[CanonicalAddress]:
    """Return `dynamic_indirect` precedents of `host_cell`, minus `exclude`.

    Argument cells of `INDIRECT(ref)` are also tagged `dynamic_indirect` at
    graph-build time. Pass those addresses as `exclude` so the remaining set
    is the resolved target, not the address-text producer.
    """
    skipped = set(exclude)
    found: list[CanonicalAddress] = []
    for dep in graph.get_dependencies(host_cell):
        addr = _canonical(str(dep))
        if addr in skipped:
            continue
        provenance = graph.get_edge_attrs(host_cell, dep).provenance
        if provenance is None:
            continue
        if DependencyCause.dynamic_indirect in provenance.causes:
            found.append(addr)
    return found


def _producer_hits(
    host: BoundSeries,
    producer: BoundSeries,
    graph: DependencyGraph,
    cells: Sequence[CanonicalAddress] | None = None,
) -> list[set[int]]:
    """Return catalog indices of `producer` read by each host member in `cells`."""
    hits: list[set[int]] = []
    for cell in host.cells if cells is None else cells:
        found: set[int] = set()
        for dep in graph.get_dependencies(cell):
            idx = producer.index_of(_canonical(str(dep)))
            if idx is not None:
                found.add(idx)
        hits.append(found)
    return hits


def _axis_sets(
    hits: Sequence[set[int]],
    width: int,
    *,
    axis: Literal["row", "col"],
) -> list[set[int]]:
    sets: list[set[int]] = []
    for indices in hits:
        coords: set[int] = set()
        for idx in indices:
            row, col = divmod(idx, width)
            coords.add(row if axis == "row" else col)
        sets.append(coords)
    return sets


def _classify_axis(
    sets: Sequence[set[int]],
    full_size: int,
    *,
    host: BoundSeries,
    producer: BoundSeries,
    axis: str,
) -> AxisAccess:
    if all(len(item) == 0 for item in sets):
        raise InvertedTreeExportError(
            f"series {host.series_id!r} cell {host.cells[0]}: "
            f"no resolved edges into producer {producer.series_id!r} on {axis}"
        )
    if any(len(item) == 0 for item in sets):
        empty_at = next(index for index, item in enumerate(sets) if not item)
        raise InvertedTreeExportError(
            f"series {host.series_id!r} cell {host.cells[empty_at]}: "
            f"producer {producer.series_id!r} {axis} is not static, dynamic, or whole"
        )
    mins = [min(item) for item in sets]
    origin_pairs = list(enumerate(mins))
    origin_values = set(mins)
    origin = (
        (0, next(iter(origin_values))) if len(origin_values) == 1 else fit_affine_map(origin_pairs)
    )
    if full_size > 0 and all(item == set(range(full_size)) for item in sets):
        return AxisAccess("whole", 0, 0)
    if all(len(item) == 1 for item in sets):
        if origin is None:
            raise InvertedTreeExportError(
                f"series {host.series_id!r} cell {host.cells[0]}: "
                f"producer {producer.series_id!r} {axis} is not an affine static map"
            )
        return AxisAccess("static", coeff=origin[0], offset=origin[1])
    if origin is None:
        raise InvertedTreeExportError(
            f"series {host.series_id!r} cell {host.cells[0]}: "
            f"producer {producer.series_id!r} {axis} window is not an affine origin"
        )
    return AxisAccess("dynamic", coeff=origin[0], offset=origin[1])


def _is_runtime_selector(node: AstNode | None) -> bool:
    return not (node is None or isinstance(node, NumberNode))


def refine_access_with_selectors(
    access: AccessFunction,
    *,
    row_arg: AstNode | None = None,
    col_arg: AstNode | None = None,
) -> AccessFunction:
    """Upgrade a whole/static axis to `dynamic` when the AST selects at runtime."""
    row = access.row
    col = access.col
    if (
        row.kind == "whole"
        and row_arg is not None
        or _is_runtime_selector(row_arg)
        and row.kind == "static"
    ):
        row = AxisAccess("dynamic", 0, 0)
    if (
        col.kind == "whole"
        and col_arg is not None
        or _is_runtime_selector(col_arg)
        and col.kind == "static"
    ):
        col = AxisAccess("dynamic", 0, 0)
    return AccessFunction(
        host_id=access.host_id,
        producer_id=access.producer_id,
        row=row,
        col=col,
        width=access.width,
    )


def classify_producer_access(
    host: BoundSeries,
    producer: BoundSeries,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    cells: Sequence[CanonicalAddress] | None = None,
) -> AccessFunction:
    """Classify each producer axis from resolved host→producer edge sets.

    One access function per `(statement, producer)`: pass the statement's
    `cells` so members of another statement (a recurrence after an `INDEX`
    seed, say) do not contribute empty hit-sets. Defaults to every member.

    Raises:
        InvertedTreeExportError: An axis is not `static`, `dynamic`, or `whole`.
            The message names the host cell and producer.
    """
    del catalog
    width = producer.block_width
    hits = _producer_hits(host, producer, graph, cells)
    n_rows = max(1, (len(producer.cells) + width - 1) // width)
    n_cols = width
    row = _classify_axis(
        _axis_sets(hits, width, axis="row"), n_rows, host=host, producer=producer, axis="row"
    )
    col = _classify_axis(
        _axis_sets(hits, width, axis="col"), n_cols, host=host, producer=producer, axis="col"
    )
    return AccessFunction(
        host_id=host.series_id,
        producer_id=producer.series_id,
        row=row,
        col=col,
        width=width,
    )


def _formula_ast(graph: DependencyGraph, address: CanonicalAddress) -> AstNode:
    node = graph.get_node(address)
    ast = getattr(node, "formula_ast", None) if node is not None else None
    if ast is None:
        raise InvertedTreeExportError(
            f"bound cell {address} has no formula AST (cannot classify cell-ref access)"
        )
    return ast


def _iter_direct_cell_refs(node: AstNode) -> Iterator[CellRefNode]:
    """Yield each `CellRefNode` in `node` (not range endpoints)."""
    match node:
        case CellRefNode():
            yield node
        case BinaryOpNode(left=left, right=right):
            yield from _iter_direct_cell_refs(left)
            yield from _iter_direct_cell_refs(right)
        case UnaryOpNode(operand=operand):
            yield from _iter_direct_cell_refs(operand)
        case FunctionCallNode(args=args):
            for arg in args:
                yield from _iter_direct_cell_refs(arg)
        case _:
            return


def _cell_ref_walk_slot(ast: AstNode, ref: CellRefNode) -> int:
    for index, node in enumerate(_iter_direct_cell_refs(ast)):
        if node is ref:
            return index
    for index, node in enumerate(_iter_direct_cell_refs(ast)):
        if node == ref:
            return index
    raise InvertedTreeExportError("cell reference is not a leaf of the host formula")


def catalog_index_affine(access: AccessFunction) -> tuple[int, int]:
    """Return `(coeff, offset)` of the flat catalog index `a*i + b`.

    A cell-ref site is one catalog slot, stored as a static column affine
    with a static zero row so `row.coeff * width + col.coeff` reconstructs
    the catalog map.
    """
    if access.row.kind != "static" or access.col.kind != "static":
        raise InvertedTreeExportError(
            f"series {access.host_id!r}: producer {access.producer_id!r} "
            "cell-ref access is not static on both axes"
        )
    return (
        access.row.coeff * access.width + access.col.coeff,
        access.row.offset * access.width + access.col.offset,
    )


def _access_from_catalog_pairs(
    host: BoundSeries,
    producer: BoundSeries,
    pairs: Sequence[tuple[int, int]],
) -> AccessFunction:
    if not pairs:
        raise InvertedTreeExportError(
            f"series {host.series_id!r}: no resolved cell-ref edges into {producer.series_id!r}"
        )
    fitted = fit_affine_map(pairs)
    if fitted is None:
        raise InvertedTreeExportError(
            f"series {host.series_id!r}: producer {producer.series_id!r} "
            "cell-ref site is not an affine static map"
        )
    coeff, offset = fitted
    return AccessFunction(
        host_id=host.series_id,
        producer_id=producer.series_id,
        row=AxisAccess("static", 0, 0),
        col=AxisAccess("static", coeff, offset),
        width=producer.block_width,
    )


def _catalog_pairs_for_slot(
    host: BoundSeries,
    producer: BoundSeries,
    graph: DependencyGraph,
    cells: Sequence[CanonicalAddress],
    slot: int,
) -> list[tuple[int, int]]:
    pairs: list[tuple[int, int]] = []
    for cell in cells:
        refs = list(_iter_direct_cell_refs(_formula_ast(graph, cell)))
        if slot >= len(refs):
            continue
        host_index = host.index_of(cell)
        prod_index = producer.index_of(as_canonical(resolve_cell_ref(refs[slot], cell)))
        if host_index is None or prod_index is None:
            continue
        pairs.append((host_index, prod_index))
    return pairs


def classify_cell_ref_accesses(
    host: BoundSeries,
    producer: BoundSeries,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    cells: Sequence[CanonicalAddress] | None = None,
) -> tuple[AccessFunction, ...]:
    """Return one access function per `CellRefNode` site of `producer`.

    Sites are matched by walk order. A mixed relative and absolute read of
    the same producer is two accesses, not one merged edge set.

    Raises:
        InvertedTreeExportError: A site is not an affine static catalog map.
    """
    del catalog
    members = tuple(host.cells if cells is None else cells)
    if not members:
        return ()
    found: list[AccessFunction] = []
    for slot, ref in enumerate(_iter_direct_cell_refs(_formula_ast(graph, members[0]))):
        address = as_canonical(resolve_cell_ref(ref, members[0]))
        if producer.index_of(address) is None:
            continue
        found.append(
            _access_from_catalog_pairs(
                host,
                producer,
                _catalog_pairs_for_slot(host, producer, graph, members, slot),
            )
        )
    return tuple(found)


def classify_cell_ref_access(
    host: BoundSeries,
    producer: BoundSeries,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    host_cell: CanonicalAddress,
    ref: CellRefNode,
    cells: Sequence[CanonicalAddress] | None = None,
) -> AccessFunction:
    """Classify the catalog-index map of one `CellRefNode` site.

    Raises:
        InvertedTreeExportError: The site is not an affine static catalog map.
    """
    del catalog
    members = tuple(host.cells if cells is None else cells)
    slot = _cell_ref_walk_slot(_formula_ast(graph, host_cell), ref)
    return _access_from_catalog_pairs(
        host,
        producer,
        _catalog_pairs_for_slot(host, producer, graph, members, slot),
    )


def _has_schedule_axis(series: BoundSeries, catalog: SeriesCatalog) -> bool:
    fields = preferred_fields(series, catalog)
    return fields is not None and "TIME_PERIOD" in fields


def overlapping_schedule_peer(
    host: BoundSeries,
    producer: BoundSeries,
    catalog: SeriesCatalog,
) -> bool:
    """True when `host` and `producer` share a key domain and a schedule coord.

    An overlapping peer is a lagged or aligned series, not a year-0 seed.
    """
    host_fields = preferred_fields(host, catalog)
    prod_fields = preferred_fields(producer, catalog)
    if host_fields is None or host_fields != prod_fields:
        return False
    host_coords = {schedule_coord(cell, catalog) for cell in host.cells}
    prod_coords = {schedule_coord(cell, catalog) for cell in producer.cells}
    return bool(host_coords & prod_coords)


def is_seed_access(
    host: BoundSeries,
    producer: BoundSeries,
    producer_cell: CanonicalAddress,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    *,
    delta: int,
) -> bool:
    """True when this read is a seed: no schedule axis, or host axis ± 1.

    A seed is a relative read of a producer with no schedule axis (scalar
    or unkeyed), or of one at schedule-axis coordinate `host ± 1` that is
    not an overlapping keyed peer. Everything else is an aligned read.
    """
    if producer.series_id == host.series_id:
        return False
    if overlapping_schedule_peer(host, producer, catalog):
        return False
    if not _has_schedule_axis(producer, catalog):
        return producer.is_scalar
    if not _has_schedule_axis(host, catalog):
        return False
    return (
        schedule_axis_coord(producer_cell, catalog)
        == schedule_axis_coord(host_cell, catalog) + delta
    )


def unique_seed_or_none(
    host: BoundSeries,
    host_cell: CanonicalAddress,
    catalog: SeriesCatalog,
    matched: Sequence[CanonicalAddress],
    *,
    delta: int,
) -> CanonicalAddress | None:
    """Return the unique seed, or `None` when several unkeyed scalars match.

    Raises:
        InvertedTreeExportError: Several candidate seeds at `host ± 1`.
    """
    matched = list(dict.fromkeys(matched))
    if len(matched) == 1:
        return matched[0]
    if not matched:
        return None
    adjacent: list[CanonicalAddress] = []
    for address in matched:
        owner = catalog.series_for(address)
        if owner is None or not _has_schedule_axis(owner, catalog):
            continue
        if schedule_axis_coord(address, catalog) == schedule_axis_coord(host_cell, catalog) + delta:
            adjacent.append(address)
    if len(adjacent) > 1:
        raise InvertedTreeExportError(
            f"series {host.series_id!r} cell {host_cell}: "
            f"ambiguous seed candidates {tuple(matched)}"
        )
    if len(adjacent) == 1:
        return adjacent[0]
    return None


def seed_address(
    host: BoundSeries,
    index: int,
    catalog: SeriesCatalog,
    graph: DependencyGraph | None,
    *,
    delta: int,
) -> CanonicalAddress | None:
    """Return the unique seed/terminal cell at `host` index, if any.

    Raises:
        InvertedTreeExportError: Several candidate seeds at `host ± 1`.
    """
    if graph is None or index < 0 or index >= len(host.cells):
        return None
    host_cell = host.cells[index]
    matched: list[CanonicalAddress] = []
    for dep in graph.get_dependencies(host_cell):
        address = _canonical(str(dep))
        owner = catalog.series_for(address)
        if owner is None or owner.series_id == host.series_id:
            continue
        if not is_seed_access(host, owner, address, host_cell, catalog, delta=delta):
            continue
        matched.append(address)
    return unique_seed_or_none(host, host_cell, catalog, matched, delta=delta)
