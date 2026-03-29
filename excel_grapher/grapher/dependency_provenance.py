from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum


class DependencyCause(StrEnum):
    """How a dependency edge arises from a formula."""

    direct_ref = "direct_ref"
    static_range = "static_range"
    dynamic_offset = "dynamic_offset"
    dynamic_indirect = "dynamic_indirect"


@dataclass(frozen=True, slots=True)
class EdgeProvenance:
    """Metadata for a single directed edge, possibly from multiple mechanisms in one formula."""

    causes: frozenset[DependencyCause]
    direct_sites_formula: tuple[tuple[int, int], ...] = ()
    direct_sites_normalized: tuple[tuple[int, int], ...] = ()

    @staticmethod
    def empty() -> EdgeProvenance:
        return EdgeProvenance(causes=frozenset())

    def merge(self, other: EdgeProvenance) -> EdgeProvenance:
        return EdgeProvenance(
            causes=self.causes | other.causes,
            direct_sites_formula=tuple(
                sorted(set(self.direct_sites_formula) | set(other.direct_sites_formula))
            ),
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
