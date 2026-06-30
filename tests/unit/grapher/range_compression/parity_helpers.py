"""Shared parity helpers for TACO index tests."""

from __future__ import annotations

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey
from excel_grapher.grapher.range_compression import (
    TacoIndex,
    materialize_dependents,
    materialize_precedents,
)


def assert_taco_parity(graph: DependencyGraph, index: TacoIndex) -> None:
    """Assert materialized TACO queries match the canonical dependency graph."""
    for key in graph:
        assert materialize_dependents(index, key) == set(graph.get_dependents(key)), (
            f"dependents mismatch at {key}"
        )
        assert materialize_precedents(index, key) == set(graph.get_dependencies(key)), (
            f"precedents mismatch at {key}"
        )


def assert_taco_parity_subset(graph: DependencyGraph, index: TacoIndex, keys: set[NodeKey]) -> None:
    """Assert parity for an explicit subset of nodes."""
    for key in keys:
        assert materialize_dependents(index, key) == set(graph.get_dependents(key)), (
            f"dependents mismatch at {key}"
        )
        assert materialize_precedents(index, key) == set(graph.get_dependencies(key)), (
            f"precedents mismatch at {key}"
        )
