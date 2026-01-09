"""
excel_grapher: Build and analyze dependency graphs from Excel workbooks.

This package intentionally keeps the public API small and stable.
"""

from .builder import create_dependency_graph
from .export import to_graphviz, to_mermaid, to_networkx
from .extractor import discover_formula_cells_in_rows
from .graph import CycleError, CycleReport, DependencyGraph, NodeHook
from .guard import And, Compare, GuardExpr, Literal, Not, Or
from .guard import CellRef as GuardCellRef
from .node import Node, NodeKey, ValueType
from .parser import format_cell_key, format_key, needs_quoting
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
    "format_cell_key",
    "format_key",
    "needs_quoting",
    # Extractor module
    "discover_formula_cells_in_rows",
]

