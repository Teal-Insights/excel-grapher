"""Non-overlapping packings of compressible candidates."""

from __future__ import annotations

import heapq
from collections import Counter
from dataclasses import dataclass

from .candidates import CompressibleCandidate
from .config import SimilarityCompressionConfig

__all__ = ["Packing", "enumerate_packings", "packing_sort_key"]


@dataclass(frozen=True)
class Packing:
    """A non-overlapping set of compressible groups applied together."""

    groups: tuple[CompressibleCandidate, ...]

    @property
    def total_reduction(self) -> int:
        """Sum of per-group node removals."""
        return sum(group.size_reduction for group in self.groups)

    @property
    def member_nodes(self) -> frozenset[str]:
        """All node keys claimed by this packing."""
        nodes: set[str] = set()
        for group in self.groups:
            nodes.update(group.members)
        return frozenset(nodes)


def packing_sort_key(packing: Packing) -> tuple[int, int, int, tuple[str, ...]]:
    """Return a descending sort key for packings.

    Primary: total node reduction. Secondary: fewer groups. Tertiary: more
    candidates sharing the same member-count profile (parallel-family proxy).
    """
    member_counts = Counter(len(group.members) for group in packing.groups)
    parallel_family_count = max(member_counts.values()) if member_counts else 0
    return (
        packing.total_reduction,
        -len(packing.groups),
        parallel_family_count,
        tuple(sorted(group.root for group in packing.groups)),
    )


def _add_packing(
    heap: list[tuple[tuple[int, int, int, tuple[str, ...]], int, Packing]],
    packing: Packing,
    *,
    max_packings: int,
    sequence: int,
) -> int:
    key = packing_sort_key(packing)
    entry = (key, sequence, packing)
    if len(heap) < max_packings:
        heapq.heappush(heap, entry)
    elif key > heap[0][0]:
        heapq.heapreplace(heap, entry)
    return sequence + 1


def enumerate_packings(
    candidates: tuple[CompressibleCandidate, ...] | list[CompressibleCandidate],
    *,
    config: SimilarityCompressionConfig | None = None,
) -> tuple[Packing, ...]:
    """Enumerate top non-overlapping packings by node reduction.

    Uses branch-and-bound depth-first search with an upper-bound prune. Keeps at
    most ``config.top_n_packings`` packings by ``packing_sort_key``.

    Args:
        candidates: Compressible subgraph candidates (typically from
            ``enumerate_compressible_candidates``).
        config: Caps the number of returned packings.

    Returns:
        Packings sorted best-first by ``packing_sort_key``.
    """
    cfg = config or SimilarityCompressionConfig()
    ordered = sorted(
        candidates,
        key=lambda c: (-c.size_reduction, c.root),
    )
    if not ordered:
        return ()
    reductions = [candidate.size_reduction for candidate in ordered]
    suffix_sums = [0] * (len(reductions) + 1)
    for index in range(len(reductions) - 1, -1, -1):
        suffix_sums[index] = suffix_sums[index + 1] + reductions[index]

    heap: list[tuple[tuple[int, int, int, tuple[str, ...]], int, Packing]] = []
    sequence = 0
    worst_kept = (-1, 0, 0, ())

    def search(
        index: int,
        chosen: list[CompressibleCandidate],
        used: frozenset[str],
        reduction: int,
    ) -> None:
        nonlocal sequence, worst_kept
        if index == len(ordered):
            if chosen:
                sequence = _add_packing(
                    heap,
                    Packing(groups=tuple(chosen)),
                    max_packings=cfg.top_n_packings,
                    sequence=sequence,
                )
                if len(heap) == cfg.top_n_packings:
                    worst_kept = heap[0][0]
            return

        remaining_upper = reduction + suffix_sums[index]
        if heap and remaining_upper < worst_kept[0]:
            return

        search(index + 1, chosen, used, reduction)

        candidate = ordered[index]
        if candidate.members.isdisjoint(used):
            search(
                index + 1,
                [*chosen, candidate],
                used | candidate.members,
                reduction + candidate.size_reduction,
            )

    search(0, [], frozenset(), 0)
    ranked = sorted(heap, key=lambda entry: entry[0], reverse=True)
    return tuple(entry[2] for entry in ranked)
