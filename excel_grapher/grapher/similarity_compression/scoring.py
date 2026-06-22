"""Score packings by parallel-family embedding tightness."""

from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass

from ..graph import DependencyGraph, NodeKey
from .canonicalize import canonicalize_for_embedding
from .config import SimilarityCompressionConfig
from .embedding import EmbeddingProvider, embed_texts
from .packings import Packing, packing_sort_key
from .signatures import StructuralSignature, structural_signature
from .simulate import SimulatedCollapse, simulate_packing

__all__ = [
    "PackingScore",
    "cluster_roots_by_signature",
    "cosine_distance",
    "mean_cluster_distance",
    "score_packing",
    "score_packings",
    "select_best_packing",
]


@dataclass(frozen=True)
class PackingScore:
    """Similarity-aware score for one packing."""

    packing: Packing
    total_reduction: int
    mean_cluster_distance: float
    singleton_cluster_fraction: float
    final_score: float


def cosine_distance(left: list[float], right: list[float]) -> float:
    """Return cosine distance ``1 - cosine_similarity``."""
    dot = sum(a * b for a, b in zip(left, right, strict=True))
    left_norm = math.sqrt(sum(a * a for a in left))
    right_norm = math.sqrt(sum(b * b for b in right))
    if left_norm == 0 or right_norm == 0:
        return 1.0
    similarity = dot / (left_norm * right_norm)
    return 1.0 - max(-1.0, min(1.0, similarity))


def cluster_roots_by_signature(
    roots: tuple[NodeKey, ...],
    collapsed_roots: dict[str, str],
    graph: DependencyGraph,
) -> dict[StructuralSignature, tuple[str, ...]]:
    """Bucket collapsed roots by coarse structural signature."""
    clusters: dict[StructuralSignature, list[str]] = defaultdict(list)
    for root in roots:
        formula = collapsed_roots.get(root)
        if formula is None:
            continue
        signature = structural_signature(root, formula, graph)
        clusters[signature].append(root)
    return {signature: tuple(sorted(keys)) for signature, keys in clusters.items()}


def mean_cluster_distance(
    clusters: dict[StructuralSignature, tuple[str, ...]],
    embeddings: dict[str, list[float]],
) -> tuple[float, float]:
    """Return mean intra-cluster distance and singleton-cluster fraction."""
    if not clusters:
        return 0.0, 0.0

    cluster_distances: list[float] = []
    singleton_clusters = 0
    for roots in clusters.values():
        if len(roots) < 2:
            singleton_clusters += 1
            continue
        vectors = [embeddings[root] for root in roots]
        pairwise = [
            cosine_distance(vectors[i], vectors[j])
            for i in range(len(vectors))
            for j in range(i + 1, len(vectors))
        ]
        cluster_distances.append(sum(pairwise) / len(pairwise))

    mean_distance = sum(cluster_distances) / len(cluster_distances) if cluster_distances else 0.0
    singleton_fraction = singleton_clusters / len(clusters)
    return mean_distance, singleton_fraction


def score_packing(
    packing: Packing,
    collapsed_roots: dict[str, str],
    graph: DependencyGraph,
    embeddings: dict[str, list[float]],
    *,
    config: SimilarityCompressionConfig | None = None,
) -> PackingScore:
    """Score one packing using reduction and embedding-cluster tightness."""
    cfg = config or SimilarityCompressionConfig()
    roots = tuple(group.root for group in packing.groups)
    clusters = cluster_roots_by_signature(roots, collapsed_roots, graph)
    mean_distance, singleton_fraction = mean_cluster_distance(clusters, embeddings)
    final_score = (
        cfg.alpha * packing.total_reduction
        - cfg.beta * mean_distance
        - cfg.gamma * singleton_fraction
    )
    return PackingScore(
        packing=packing,
        total_reduction=packing.total_reduction,
        mean_cluster_distance=mean_distance,
        singleton_cluster_fraction=singleton_fraction,
        final_score=final_score,
    )


def score_packings(
    graph: DependencyGraph,
    packings: tuple[Packing, ...],
    *,
    provider: EmbeddingProvider,
    config: SimilarityCompressionConfig | None = None,
    preserve: set[NodeKey] | None = None,
) -> list[tuple[PackingScore, SimulatedCollapse]]:
    """Simulate and score each packing, sharing embedding work by canonical blob."""
    cfg = config or SimilarityCompressionConfig()
    simulations = [simulate_packing(graph, packing, preserve=preserve) for packing in packings]
    canonical_by_root: dict[str, str] = {}
    for simulation in simulations:
        for root, formula in simulation.collapsed_roots.items():
            canonical_by_root[root] = canonicalize_for_embedding(root, formula, graph)
    unique_texts = sorted(set(canonical_by_root.values()))
    vectors_by_text = embed_texts(unique_texts, provider)
    embeddings_by_root = {root: vectors_by_text[text] for root, text in canonical_by_root.items()}

    scored: list[tuple[PackingScore, SimulatedCollapse]] = []
    for packing, simulation in zip(packings, simulations, strict=True):
        packing_embeddings = {root: embeddings_by_root[root] for root in simulation.collapsed_roots}
        score = score_packing(
            packing,
            simulation.collapsed_roots,
            graph,
            packing_embeddings,
            config=cfg,
        )
        scored.append((score, simulation))
    return scored


def select_best_packing(
    graph: DependencyGraph,
    packings: tuple[Packing, ...],
    *,
    provider: EmbeddingProvider,
    config: SimilarityCompressionConfig | None = None,
    preserve: set[NodeKey] | None = None,
) -> tuple[PackingScore, SimulatedCollapse]:
    """Pick the highest-scoring packing, with reduction-only fallback on flat scores."""
    if not packings:
        raise ValueError("select_best_packing requires at least one packing")

    cfg = config or SimilarityCompressionConfig()
    scored = score_packings(
        graph,
        packings,
        provider=provider,
        config=cfg,
        preserve=preserve,
    )
    scored.sort(key=lambda item: item[0].final_score, reverse=True)
    best_score, best_simulation = scored[0]
    if len(scored) > 1 and cfg.fallback_to_optimal:
        spread = scored[0][0].final_score - scored[-1][0].final_score
        if spread < cfg.score_flatness_epsilon:
            by_reduction = max(scored, key=lambda item: packing_sort_key(item[0].packing))
            return by_reduction
    return best_score, best_simulation
