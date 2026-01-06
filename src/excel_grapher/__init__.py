"""
excel_grapher: Build and analyze dependency graphs from Excel workbooks.

This package intentionally keeps the public API small and stable.
"""

from .builder import create_dependency_graph
from .export import to_graphviz, to_mermaid, to_networkx
from .graph import CycleError, CycleReport, DependencyGraph, NodeHook
from .guard import And, CellRef as GuardCellRef, Compare, GuardExpr, Literal, Not, Or
from .node import Node, NodeKey, ValueType
from .validation import ValidationResult, WorkbookCalcSettings, get_calc_settings, validate_graph

__all__ = [
    "create_dependency_graph",
    "DependencyGraph",
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
]

