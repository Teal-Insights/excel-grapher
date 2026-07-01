"""Optional TACO-style range-pattern compression for dependency graphs."""

from __future__ import annotations

from .build import build_taco_index
from .config import TacoBuildConfig, input_keys_from_graph
from .index import TacoIndex
from .materialize import materialize_dependents, materialize_precedents
from .types import CompressedEdge, PatternKind, PatternMeta, RangeRef, SingleEdge

__all__ = [
    "CompressedEdge",
    "PatternKind",
    "PatternMeta",
    "RangeRef",
    "SingleEdge",
    "TacoBuildConfig",
    "TacoIndex",
    "build_taco_index",
    "input_keys_from_graph",
    "materialize_dependents",
    "materialize_precedents",
]
