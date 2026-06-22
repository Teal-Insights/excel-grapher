"""End-to-end similarity-aware packing selection."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.core.address_keys import normalize_key

from ..graph import DependencyGraph, NodeKey
from .candidates import enumerate_compressible_candidates
from .config import SimilarityCompressionConfig
from .embedding import EmbeddingProvider, MockEmbeddingProvider
from .packings import Packing, enumerate_packings
from .scoring import PackingScore, select_best_packing
from .simulate import SimulatedCollapse

__all__ = ["SimilaritySelectionResult", "select_similarity_projection"]


@dataclass(frozen=True)
class SimilaritySelectionResult:
    """Best packing chosen by similarity-aware scoring."""

    packing: Packing
    simulation: SimulatedCollapse
    score: PackingScore


def select_similarity_projection(
    graph: DependencyGraph,
    *,
    preserve: set[NodeKey] | None = None,
    config: SimilarityCompressionConfig | None = None,
    provider: EmbeddingProvider | None = None,
) -> SimilaritySelectionResult:
    """Enumerate, score, and return the best similarity-aware packing.

    Args:
        graph: Canonical dependency graph with provenance.
        preserve: Node keys that must not be inlined away.
        config: Search limits and scoring weights.
        provider: Embedding backend (defaults to ``MockEmbeddingProvider``).

    Returns:
        Selected packing, its simulation, and the score breakdown.
    """
    cfg = config or SimilarityCompressionConfig()
    embedder = provider or MockEmbeddingProvider()
    if preserve is None:
        preserve_set: set[NodeKey] | None = None
    else:
        preserve_set = {normalize_key(key) for key in preserve}

    candidates = enumerate_compressible_candidates(
        graph,
        preserve=preserve_set,
        config=cfg,
    )
    packings = enumerate_packings(candidates, config=cfg)
    if not packings:
        raise ValueError("No compressible packings found for similarity-aware compression")

    score, simulation = select_best_packing(
        graph,
        packings,
        provider=embedder,
        config=cfg,
        preserve=preserve_set,
    )
    return SimilaritySelectionResult(
        packing=score.packing,
        simulation=simulation,
        score=score,
    )
