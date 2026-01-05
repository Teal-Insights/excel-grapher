"""
excel_grapher: Build and analyze dependency graphs from Excel workbooks.

This package intentionally keeps the public API small and stable.
"""

from .builder import create_dependency_graph
from .export import to_graphviz, to_mermaid, to_networkx
from .graph import DependencyGraph, NodeHook
from .node import Node, NodeKey, ValueType
from .validation import ValidationResult, validate_graph

__all__ = [
    "create_dependency_graph",
    "DependencyGraph",
    "NodeHook",
    "Node",
    "NodeKey",
    "ValueType",
    "to_graphviz",
    "to_mermaid",
    "to_networkx",
    "validate_graph",
    "ValidationResult",
]

