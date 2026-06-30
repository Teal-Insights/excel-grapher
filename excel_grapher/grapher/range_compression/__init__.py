"""Optional TACO-style range-pattern compression for dependency graphs."""

from __future__ import annotations

from .build import build_taco_index
from .index import TacoIndex
from .materialize import materialize_dependents, materialize_precedents
from .types import CompressedEdge, PatternKind, PatternMeta, RangeRef, SingleEdge

__all__ = [
    "CompressedEdge",
    "PatternKind",
    "PatternMeta",
    "RangeRef",
    "SingleEdge",
    "TacoIndex",
    "build_taco_index",
    "materialize_dependents",
    "materialize_precedents",
]
