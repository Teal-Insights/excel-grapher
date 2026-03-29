"""
excel_grapher: Build and analyze dependency graphs from Excel workbooks.

This package intentionally keeps the public API small and stable.
"""

from excel_grapher.core.cell_types import GreaterThanCell, NotEqualCell, RealBetween, RealIntervalDomain

from .builder import create_dependency_graph
from .dependency_provenance import DependencyCause, EdgeProvenance
from .dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    DynamicRefLimits,
    FromWorkbook,
    constrain,
    infer_dynamic_index_targets,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)
from .export import to_graphviz, to_mermaid, to_networkx
from .graph import CycleError, CycleReport, DependencyGraph, NodeHook
from .guard import And, Compare, GuardExpr, Literal, Not, Or
from .guard import CellRef as GuardCellRef
from .node import Node, NodeKey, ValueType
from .parser import format_cell_key, format_key, needs_quoting
from .validation import ValidationResult, WorkbookCalcSettings, get_calc_settings, validate_graph

__all__ = [
    "create_dependency_graph",
    "DependencyCause",
    "DependencyGraph",
    "EdgeProvenance",
    "DynamicRefConfig",
    "DynamicRefError",
    "DynamicRefLimits",
    "FromWorkbook",
    "GreaterThanCell",
    "NotEqualCell",
    "RealBetween",
    "RealIntervalDomain",
    "constrain",
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
    "ValueType",
    "to_graphviz",
    "to_mermaid",
    "to_networkx",
    "validate_graph",
    "ValidationResult",
    "get_calc_settings",
    "WorkbookCalcSettings",
    "format_cell_key",
    "format_key",
    "needs_quoting"
]

