"""Similarity-aware compression for dependency graphs (issue #282)."""

from .candidates import CompressibleCandidate, enumerate_compressible_candidates
from .config import SimilarityCompressionConfig
from .packings import Packing, enumerate_packings, packing_sort_key
from .simulate import SimulatedCollapse, collapse_candidate_on_graph, simulate_packing

__all__ = [
    "CompressibleCandidate",
    "Packing",
    "SimilarityCompressionConfig",
    "SimulatedCollapse",
    "collapse_candidate_on_graph",
    "enumerate_compressible_candidates",
    "enumerate_packings",
    "packing_sort_key",
    "simulate_packing",
]
