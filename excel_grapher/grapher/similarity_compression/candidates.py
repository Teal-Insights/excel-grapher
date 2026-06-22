"""Compressible subgraph candidates for similarity-aware compression."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.core.address_keys import normalize_key

from ..compression import (
    _structural_inline_candidate,
    compression_safe_provenance,
    is_identity_transit,
    require_compression_provenance,
)
from ..graph import DependencyGraph, NodeKey
from .config import SimilarityCompressionConfig

__all__ = ["CompressibleCandidate", "enumerate_compressible_candidates"]


@dataclass(frozen=True)
class CompressibleCandidate:
    """One root and the node set that can collapse into it."""

    root: NodeKey
    members: frozenset[NodeKey]

    @property
    def size_reduction(self) -> int:
        """Number of nodes removed when this group is collapsed."""
        return len(self.members) - 1


def _internal_node_collapsible(
    graph: DependencyGraph,
    key: NodeKey,
    dependent: NodeKey,
) -> bool:
    replacement = is_identity_transit(graph, key)
    if replacement is not None:
        prov = graph.get_edge_attrs(dependent, key).provenance
        return compression_safe_provenance(prov)
    return _structural_inline_candidate(graph, key) == dependent


def _is_connected_component(graph: DependencyGraph, members: frozenset[NodeKey]) -> bool:
    if len(members) <= 1:
        return True
    start = next(iter(members))
    seen: set[NodeKey] = {start}
    stack = [start]
    member_set = set(members)
    while stack:
        node = stack.pop()
        neighbors = graph.get_dependencies(node) | graph.get_dependents(node)
        for neighbor in neighbors:
            if neighbor in member_set and neighbor not in seen:
                seen.add(neighbor)
                stack.append(neighbor)
    return seen == member_set


def _grow_members_for_root(
    graph: DependencyGraph,
    root: NodeKey,
    preserve: frozenset[NodeKey],
    *,
    require_connected: bool,
) -> frozenset[NodeKey] | None:
    members: set[NodeKey] = {root}
    changed = True
    while changed:
        changed = False
        for key in graph:
            if key in members or key in preserve:
                continue
            dependents = graph.get_dependents(key)
            if not dependents or not dependents.issubset(members):
                continue
            if len(dependents) != 1:
                continue
            dependent = next(iter(dependents))
            if not _internal_node_collapsible(graph, key, dependent):
                continue
            members.add(key)
            changed = True

    if len(members) < 2:
        return None
    frozen = frozenset(members)
    if require_connected and not _is_connected_component(graph, frozen):
        return None
    return frozen


def _prune_subset_candidates(
    candidates: list[CompressibleCandidate],
) -> list[CompressibleCandidate]:
    by_root: dict[NodeKey, list[CompressibleCandidate]] = {}
    for candidate in candidates:
        by_root.setdefault(candidate.root, []).append(candidate)

    pruned: list[CompressibleCandidate] = []
    for root_candidates in by_root.values():
        ordered = sorted(root_candidates, key=lambda c: -len(c.members))
        kept: list[CompressibleCandidate] = []
        for candidate in ordered:
            if any(candidate.members < other.members for other in kept):
                continue
            kept.append(candidate)
        pruned.extend(kept)
    return pruned


def enumerate_compressible_candidates(
    graph: DependencyGraph,
    *,
    preserve: set[NodeKey] | None = None,
    config: SimilarityCompressionConfig | None = None,
) -> tuple[CompressibleCandidate, ...]:
    """Enumerate safe compressible subgraphs with a single external root.

    Grows upstream from each potential root by repeatedly attaching nodes whose
    sole dependent is already in the group. Applies the same substitution-safety
    checks used by optimal compression.

    Args:
        graph: Canonical dependency graph with provenance.
        preserve: Node keys that must not be inlined away (defaults to targets).
        config: Search limits and connectivity requirement.

    Returns:
        Candidates sorted by ``size_reduction`` descending, capped at
        ``config.max_candidates``.
    """
    cfg = config or SimilarityCompressionConfig()
    require_compression_provenance(graph)
    if preserve is None:
        preserve_set = frozenset(graph.target_keys())
    else:
        preserve_set = frozenset(normalize_key(key) for key in preserve)

    seen: set[tuple[NodeKey, frozenset[NodeKey]]] = set()
    candidates: list[CompressibleCandidate] = []
    for root in graph:
        members = _grow_members_for_root(
            graph,
            root,
            preserve_set,
            require_connected=cfg.require_connected_component,
        )
        if members is None:
            continue
        if len(members) > cfg.max_members_per_candidate:
            continue
        dedupe_key = (root, members)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)
        candidates.append(CompressibleCandidate(root=root, members=members))

    pruned = _prune_subset_candidates(candidates)
    pruned.sort(key=lambda c: (-c.size_reduction, c.root))
    return tuple(pruned[: cfg.max_candidates])
