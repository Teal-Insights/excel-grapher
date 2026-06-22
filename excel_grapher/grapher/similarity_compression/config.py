"""Configuration for similarity-aware graph compression (issue #282)."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SimilarityCompressionConfig:
    """Tunable parameters for candidate search, packing, embedding, and scoring.

    Defaults match the Sprint 0 decisions for issue #282: prioritize parallel-family
    tightness (``beta``) over raw node reduction (``alpha``), cap combinatorial
    search, and fall back to ``OptimalCompression`` when similarity scores are flat.
    """

    max_candidates: int = 200
    top_n_packings: int = 50
    max_members_per_candidate: int = 50
    require_connected_component: bool = True
    alpha: float = 0.4
    beta: float = 0.6
    gamma: float = 0.05
    score_flatness_epsilon: float = 0.01
    embedding_model: str = "text-embedding-3-small"
    fallback_to_optimal: bool = True
