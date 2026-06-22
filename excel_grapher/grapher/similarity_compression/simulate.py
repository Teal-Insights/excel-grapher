"""Non-destructive collapse simulation for similarity-aware compression."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.core.address_keys import normalize_key

from ..compression import (
    IdentityTransitCompressionRecord,
    OptimalCompressionRecord,
    clear_identity_singleton_ref_cache,
    is_identity_transit,
    require_compression_provenance,
    snapshot_transit_node,
)
from ..graph import CycleError, DependencyGraph, NodeKey
from .candidates import CompressibleCandidate
from .packings import Packing

__all__ = ["SimulatedCollapse", "collapse_candidate_on_graph", "simulate_packing"]


@dataclass(frozen=True)
class SimulatedCollapse:
    """Result of applying one packing to a graph copy."""

    projected_graph: DependencyGraph
    record: OptimalCompressionRecord
    collapsed_roots: dict[str, str]
    packing: Packing


def _collapse_order(
    graph: DependencyGraph,
    members: frozenset[NodeKey],
    root: NodeKey,
) -> tuple[NodeKey, ...]:
    """Return internal nodes in dependency-first order for safe inlining."""
    internals = members - {root}
    if not internals:
        return ()
    try:
        ordered = [key for key in graph.evaluation_order(strict=False) if key in internals]
    except CycleError:
        ordered = []
    seen = set(ordered)
    for key in graph.keys(order="workbook", source=internals):
        if key not in seen:
            ordered.append(key)
            seen.add(key)
    return tuple(ordered)


def collapse_candidate_on_graph(
    graph: DependencyGraph,
    candidate: CompressibleCandidate,
    record: OptimalCompressionRecord,
) -> None:
    """Collapse one candidate group in place on ``graph``.

    Args:
        graph: Projected graph copy to mutate.
        candidate: Compressible group rooted at ``candidate.root``.
        record: Lineage record to populate.
    """
    clear_identity_singleton_ref_cache()
    try:
        for member in _collapse_order(graph, candidate.members, candidate.root):
            if member not in graph:
                continue
            replacement = is_identity_transit(graph, member)
            if replacement is not None:
                snapshot = snapshot_transit_node(graph, member)
                transit_record = IdentityTransitCompressionRecord()
                graph._compress_one_transit(member, replacement, record=transit_record)
                record.note_forwarding(member, replacement, snapshot)
                record.formula_rewrites.extend(transit_record.formula_rewrites)
                continue

            dependents = graph.get_dependents(member) & candidate.members
            if len(dependents) != 1:
                msg = (
                    f"Cannot collapse {member!r} in group rooted at "
                    f"{candidate.root!r}: expected one in-group dependent, "
                    f"got {sorted(dependents)!r}."
                )
                raise ValueError(msg)
            dependent = next(iter(dependents))
            snapshot = snapshot_transit_node(graph, member)
            graph._inline_one_node(member, dependent, record=record)
            record.note_inline(member, dependent, snapshot)
    finally:
        clear_identity_singleton_ref_cache()


def simulate_packing(
    graph: DependencyGraph,
    packing: Packing,
    *,
    preserve: set[NodeKey] | None = None,
) -> SimulatedCollapse:
    """Apply a packing to a graph copy and record collapsed root formulas.

    The canonical ``graph`` is not mutated.

    Args:
        graph: Canonical dependency graph with provenance.
        packing: Non-overlapping compressible groups to apply together.
        preserve: Reserved for future preserve checks during simulation.

    Returns:
        Projected graph, compression lineage, and per-root collapsed formulas.
    """
    del preserve  # Sprint 2 applies explicit packings; preserve enforced at candidate search.
    require_compression_provenance(graph)
    projected = graph.copy()
    record = OptimalCompressionRecord()
    collapsed_roots: dict[str, str] = {}

    for candidate in packing.groups:
        collapse_candidate_on_graph(projected, candidate, record)
        root_node = projected.get_node(candidate.root)
        if root_node is not None and root_node.normalized_formula is not None:
            collapsed_roots[normalize_key(candidate.root)] = root_node.normalized_formula

    return SimulatedCollapse(
        projected_graph=projected,
        record=record,
        collapsed_roots=collapsed_roots,
        packing=packing,
    )
