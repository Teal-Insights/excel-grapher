"""excel_grapher: Build and analyze dependency graphs from Excel workbooks.

This package intentionally keeps the public API small and stable.
"""

from excel_grapher.core.cell_types import (
    GreaterThanCell,
    NotEqualCell,
    RealBetween,
    RealIntervalDomain,
)

from .blank_ranges import normalize_blank_range_specs
from .builder import create_dependency_graph, list_dynamic_ref_constraint_candidates
from .cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    CacheValidationPolicy,
    build_graph_cache_meta,
    build_graph_cache_meta_portable,
    save_graph_cache,
    try_load_graph_cache,
)
from .dependency_provenance import DependencyCause, EdgeProvenance
from .dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    DynamicRefLimits,
    DynamicRefTraceEvent,
    DynamicRefTraceFn,
    FromWorkbook,
    infer_dynamic_index_targets,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
    trace_dynamic_refs,
)
from .export import (
    LightweightVizLocalEdges,
    LightweightVizModule,
    LightweightVizModuleEdge,
    LightweightVizNodeColumns,
    LightweightVizPayload,
    LightweightVizStats,
    select_path_induced_subgraph,
    to_graphviz,
    to_mermaid,
    to_networkx,
    write_lightweight_viz_data,
    write_web_viz_html,
)
from .graph import CycleError, CycleReport, DependencyGraph, GraphReadView, NodeHook
from .guard import And, Compare, GuardExpr, Literal, Not, Or
from .guard import CellRef as GuardCellRef
from .node import Node, NodeKey
from .parser import format_cell_key, format_key, needs_quoting
from .preparsed_formulas import warm_preparsed_formulas
from .range_compression import (
    CodegenPlan,
    TacoBuildConfig,
    TacoIndex,
    build_codegen_plan,
    build_codegen_taco_index,
    build_taco_index,
    codegen_boundary_keys,
    input_keys_from_graph,
    input_keys_from_ranges,
    setter_keys_from_bindings,
)
from .validation import ValidationResult, WorkbookCalcSettings, get_calc_settings, validate_graph

__all__ = [
    "create_dependency_graph",
    "build_taco_index",
    "build_codegen_taco_index",
    "build_codegen_plan",
    "CodegenPlan",
    "codegen_boundary_keys",
    "input_keys_from_graph",
    "input_keys_from_ranges",
    "setter_keys_from_bindings",
    "TacoBuildConfig",
    "TacoIndex",
    "normalize_blank_range_specs",
    "list_dynamic_ref_constraint_candidates",
    "GRAPH_CACHE_SCHEMA_VERSION",
    "build_graph_cache_meta",
    "build_graph_cache_meta_portable",
    "CacheValidationPolicy",
    "save_graph_cache",
    "try_load_graph_cache",
    "warm_preparsed_formulas",
    "DependencyCause",
    "DependencyGraph",
    "GraphReadView",
    "EdgeProvenance",
    "DynamicRefConfig",
    "DynamicRefError",
    "DynamicRefLimits",
    "DynamicRefTraceEvent",
    "DynamicRefTraceFn",
    "FromWorkbook",
    "trace_dynamic_refs",
    "GreaterThanCell",
    "NotEqualCell",
    "RealBetween",
    "RealIntervalDomain",
    "infer_dynamic_index_targets",
    "infer_dynamic_indirect_targets",
    "infer_dynamic_offset_targets",
    "NodeHook",
    "CycleError",
    "CycleReport",
    "GuardExpr",
    "GuardCellRef",
    "Literal",
    "Compare",
    "Not",
    "And",
    "Or",
    "Node",
    "NodeKey",
    "LightweightVizLocalEdges",
    "LightweightVizModule",
    "LightweightVizModuleEdge",
    "LightweightVizNodeColumns",
    "LightweightVizPayload",
    "LightweightVizStats",
    "select_path_induced_subgraph",
    "to_graphviz",
    "to_mermaid",
    "to_networkx",
    "write_web_viz_html",
    "write_lightweight_viz_data",
    "validate_graph",
    "ValidationResult",
    "get_calc_settings",
    "WorkbookCalcSettings",
    "format_cell_key",
    "format_key",
    "needs_quoting",
]
