"""Similarity-aware compression for dependency graphs (issue #282)."""

from .candidates import CompressibleCandidate, enumerate_compressible_candidates
from .config import SimilarityCompressionConfig
from .packings import Packing, enumerate_packings, packing_sort_key

__all__ = [
    "CompressibleCandidate",
    "Packing",
    "SimilarityCompressionConfig",
    "enumerate_compressible_candidates",
    "enumerate_packings",
    "packing_sort_key",
]
