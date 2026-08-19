from __future__ import annotations

from dataclasses import dataclass
from enum import IntFlag, auto


class DependencyCause(IntFlag):
    """How a dependency edge arises from a formula."""

    direct_ref = auto()
    static_range = auto()
    dynamic_offset = auto()
    dynamic_indirect = auto()
    dynamic_index = auto()


@dataclass(frozen=True, slots=True)
class EdgeProvenance:
    """Metadata for a single directed edge, possibly from multiple mechanisms in one formula.

    `direct_sites_normalized` holds `[start, end)` spans of the precedent
    reference inside the dependent's `normalized_formula`. Spans against the raw
    `Node.formula` are deliberately not stored: `normalized_formula` is the
    single source for compression / projection rewriting, and the raw formula is
    an opt-in audit field that rewriting leaves untouched.
    """

    causes: DependencyCause
    direct_sites_normalized: tuple[tuple[int, int], ...] = ()

    @staticmethod
    def empty() -> EdgeProvenance:
        return EdgeProvenance(causes=DependencyCause(0))

    def merge(self, other: EdgeProvenance) -> EdgeProvenance:
        return EdgeProvenance(
            causes=self.causes | other.causes,
            direct_sites_normalized=tuple(
                sorted(set(self.direct_sites_normalized) | set(other.direct_sites_normalized))
            ),
        )


def merge_edge_provenance(
    a: EdgeProvenance | None, b: EdgeProvenance | None
) -> EdgeProvenance | None:
    if a is None:
        return b
    if b is None:
        return a
    return a.merge(b)


def merge_provenance_maps(
    maps: list[dict[str, EdgeProvenance]],
) -> dict[str, EdgeProvenance]:
    out: dict[str, EdgeProvenance] = {}
    for m in maps:
        for k, v in m.items():
            out[k] = merge_edge_provenance(out.get(k), v) or v
    return out
