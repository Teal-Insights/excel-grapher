"""Similarity-aware compression for dependency graphs (issue #282)."""

from .candidates import CompressibleCandidate, enumerate_compressible_candidates
from .canonicalize import canonicalize_for_embedding
from .config import SimilarityCompressionConfig
from .embedding import (
    EmbeddingCache,
    EmbeddingProvider,
    MockEmbeddingProvider,
    OpenAIEmbeddingProvider,
    embed_texts,
)
from .packings import Packing, enumerate_packings, packing_sort_key
from .scoring import (
    PackingScore,
    cluster_roots_by_signature,
    cosine_distance,
    score_packing,
    score_packings,
    select_best_packing,
)
from .selection import SimilaritySelectionResult, select_similarity_projection
from .signatures import StructuralSignature, structural_signature
from .simulate import SimulatedCollapse, collapse_candidate_on_graph, simulate_packing

__all__ = [
    "CompressibleCandidate",
    "EmbeddingCache",
    "EmbeddingProvider",
    "MockEmbeddingProvider",
    "OpenAIEmbeddingProvider",
    "Packing",
    "PackingScore",
    "SimilarityCompressionConfig",
    "SimilaritySelectionResult",
    "SimulatedCollapse",
    "StructuralSignature",
    "canonicalize_for_embedding",
    "cluster_roots_by_signature",
    "collapse_candidate_on_graph",
    "cosine_distance",
    "embed_texts",
    "enumerate_compressible_candidates",
    "enumerate_packings",
    "packing_sort_key",
    "score_packing",
    "score_packings",
    "select_best_packing",
    "select_similarity_projection",
    "simulate_packing",
    "structural_signature",
]
