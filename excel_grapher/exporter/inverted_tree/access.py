"""Access functions derived from the graph's resolved edges.

One function per `(host statement, producer block)`. Each producer axis is
`static` (affine in the host index), `dynamic` (candidate set + runtime
selector), or `whole` (the block itself). Anything else fails closed.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical
from excel_grapher.core.formula_ast import AstNode, NumberNode
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    fit_affine_map,
    preferred_fields,
    schedule_axis_coord,
    schedule_coord,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

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


def _producer_hits(
    host: BoundSeries,
    producer: BoundSeries,
    graph: DependencyGraph,
) -> list[set[int]]:
    """Return catalog indices of `producer` read by each host member."""
    hits: list[set[int]] = []
    for cell in host.cells:
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
) -> AccessFunction:
    """Classify each producer axis from resolved host→producer edge sets.

    Raises:
        InvertedTreeExportError: An axis is not `static`, `dynamic`, or `whole`.
            The message names the host cell and producer.
    """
    del catalog
    width = producer.block_width
    hits = _producer_hits(host, producer, graph)
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
