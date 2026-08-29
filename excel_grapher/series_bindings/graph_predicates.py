"""Shared dependency-graph membership predicates for series binding resolution."""

from __future__ import annotations

from excel_grapher.grapher.graph import DependencyGraph


def is_graph_node(graph: DependencyGraph, address: str) -> bool:
    """Return True when `address` is present in the dependency graph."""
    return address in graph


def is_graph_leaf(graph: DependencyGraph, address: str) -> bool:
    """Return True when `address` is a graph leaf (value cell with no formula)."""
    node = graph.get_node(address) if address in graph else None
    return bool(node is not None and node.is_leaf)


def is_graph_formula_node(graph: DependencyGraph, address: str) -> bool:
    """Return True when `address` is a graph node with a workbook formula."""
    node = graph.get_node(address) if address in graph else None
    return bool(node is not None and node.has_formula)
