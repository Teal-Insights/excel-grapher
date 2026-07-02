"""Optional TACO-style range-pattern compression for dependency graphs.

Groups consecutive formula cells along a **column** (fill-down) or **row** (fill-right),
classifies RR / RF / FR / FF / RR-Chain autofill, and builds a parallel compressed
index without mutating the cell-level graph.
"""

from __future__ import annotations

from .boundaries import assert_codegen_index_boundaries
from .build import build_codegen_taco_index, build_taco_index
from .codegen_plan import (
    CodegenPlan,
    CodegenUnit,
    CompressedUnit,
    SingleCellUnit,
    build_codegen_plan,
    range_ref_unit_id,
)
from .config import (
    TacoBuildConfig,
    codegen_boundary_keys,
    input_keys_from_graph,
    input_keys_from_ranges,
    setter_keys_from_bindings,
)
from .index import TacoIndex
from .materialize import materialize_dependents, materialize_precedents
from .types import CompressedEdge, PatternKind, PatternMeta, RangeRef, SingleEdge

__all__ = [
    "CodegenPlan",
    "CodegenUnit",
    "CompressedUnit",
    "SingleCellUnit",
    "CompressedEdge",
    "PatternKind",
    "PatternMeta",
    "RangeRef",
    "SingleEdge",
    "TacoBuildConfig",
    "TacoIndex",
    "assert_codegen_index_boundaries",
    "build_codegen_plan",
    "build_codegen_taco_index",
    "build_taco_index",
    "codegen_boundary_keys",
    "input_keys_from_graph",
    "input_keys_from_ranges",
    "materialize_dependents",
    "materialize_precedents",
    "range_ref_unit_id",
    "setter_keys_from_bindings",
]
