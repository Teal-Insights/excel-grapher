"""Access functions derived from the graph's resolved edges.

Lookups classify one function per `(host statement, producer block)`.
Plain cell references classify one function per formula site of that
producer — a mixed relative and absolute read is two accesses. Each
producer axis is `static` (affine in the host index), `dynamic`
(candidate set + runtime selector), or `whole` (the block itself).
Anything else fails closed.
"""

from __future__ import annotations

import warnings
from collections.abc import Sequence
from typing import TYPE_CHECKING, Literal

from excel_grapher.core.address_keys import CanonicalAddress, as_canonical
from excel_grapher.core.formula_ast import (
    CellRefNode,
    FunctionCallNode,
    resolve_cell_ref,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    preferred_fields,
    schedule_axis_coord,
    schedule_coord,
)
from excel_grapher.grapher.dependency_provenance import DependencyCause

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph

AxisKind = Literal["static", "dynamic", "whole"]


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


def offset_target_addresses(
    graph: DependencyGraph,
    host_cell: CanonicalAddress,
    *,
    exclude: Sequence[CanonicalAddress] = (),
) -> list[CanonicalAddress]:
    """Return `dynamic_offset` precedents of `host_cell`, minus `exclude`.

    `OFFSET(INDEX(range, ...), rows, cols)` writes the destination cells as
    `dynamic_offset` edges. The INDEX array remains a static range in the
    formula; pass those addresses as `exclude` so the remaining set is the
    relocated window.
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
        if DependencyCause.dynamic_offset in provenance.causes:
            found.append(addr)
    return found


def _has_schedule_axis(series: BoundSeries, catalog: SeriesCatalog) -> bool:
    fields = preferred_fields(series, catalog)
    return fields is not None and "TIME_PERIOD" in fields


def _axis_coords(series: BoundSeries, catalog: SeriesCatalog) -> set[int]:
    axis_of = catalog.schedule.axis_of
    return {axis_of[cell] for cell in series.cells if cell in axis_of}


def _producer_covers_host_fields(
    host_fields: tuple[str, ...],
    prod_fields: tuple[str, ...],
) -> bool:
    """True when producer keys equal the host's or add only outer partition keys."""
    if host_fields == prod_fields:
        return True
    host_set = frozenset(host_fields)
    prod_set = frozenset(prod_fields)
    return "TIME_PERIOD" in host_set and "TIME_PERIOD" in prod_set and host_set < prod_set


def overlapping_schedule_peer(
    host: BoundSeries,
    producer: BoundSeries,
    catalog: SeriesCatalog,
) -> bool:
    """True when `host` and `producer` share a key domain and a schedule coord.

    An overlapping peer is a lagged or aligned series, not a year-0 seed.
    A producer whose preferred fields are a superset of the host's is a
    peer when they differ only in outer partition keys and share at least
    one inner schedule-axis coordinate (`TIME_PERIOD`, #747).
    """
    host_fields = preferred_fields(host, catalog)
    prod_fields = preferred_fields(producer, catalog)
    if host_fields is None or prod_fields is None:
        return False
    if not _producer_covers_host_fields(host_fields, prod_fields):
        return False
    if host_fields == prod_fields:
        cached = catalog.schedule.coords_of
        if host.series_id in cached and producer.series_id in cached:
            return not cached[host.series_id].isdisjoint(cached[producer.series_id])
        host_coords = {schedule_coord(cell, catalog) for cell in host.cells}
        prod_coords = {schedule_coord(cell, catalog) for cell in producer.cells}
        return bool(host_coords & prod_coords)
    return bool(_axis_coords(host, catalog) & _axis_coords(producer, catalog))


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
    not an overlapping keyed peer. A richer-keyed producer that shares the
    host's schedule axis is a peer, not a seed (#747). Everything else is
    an aligned read.
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
    """Return the unique seed, or `None` when no unique candidate exists.

    Several unkeyed scalars, or several keyed cells at `host ± 1`, are not a
    unique seed. The latter emits a `UserWarning` and returns `None` rather
    than raising: `IF(cond, A+B, C)` reads multiple producers at one offset
    by construction (#745).
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
        warnings.warn(
            f"series {host.series_id!r} cell {host_cell}: "
            f"ambiguous seed candidates {tuple(matched)}",
            UserWarning,
            stacklevel=3,
        )
        return None
    if len(adjacent) == 1:
        return adjacent[0]
    return None
