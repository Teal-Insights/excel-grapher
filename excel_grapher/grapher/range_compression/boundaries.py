"""Boundary checks for TACO compression (targets, inputs, internal-only)."""

from __future__ import annotations

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

from .config import TacoBuildConfig
from .index import TacoIndex
from .types import CompressedEdge


def dependent_may_compress(
    graph: DependencyGraph,
    key: NodeKey,
    config: TacoBuildConfig,
) -> bool:
    """Return whether a formula cell may appear in a compressed dependent range."""
    node = graph.get_node(key)
    if node is None or node.is_leaf or not node.formula:
        return False
    if config.exclude_targets and node.is_target:
        return False
    if key in config.exclude_input_keys:
        return False
    if config.internal_only:
        return is_internal_node(graph, key, config)
    return True


def precedent_may_compress(
    graph: DependencyGraph,
    key: NodeKey,
    config: TacoBuildConfig,
) -> bool:
    """Return whether a precedent cell may appear in a compressed precedent range."""
    if key in config.exclude_input_keys:
        return False
    if not config.internal_only:
        return True
    return is_internal_node(graph, key, config)


def range_keys_may_compress_as_dependents(
    graph: DependencyGraph,
    keys: list[NodeKey],
    config: TacoBuildConfig,
) -> bool:
    """Return whether every cell in `keys` may compress on the dependent side."""
    return bool(keys) and all(dependent_may_compress(graph, k, config) for k in keys)


def range_ref_dependents_may_compress(
    graph: DependencyGraph,
    range_ref,
    config: TacoBuildConfig,
) -> bool:
    """Return whether every cell in a dependent `RangeRef` may compress."""
    return range_keys_may_compress_as_dependents(graph, list(range_ref.cell_keys()), config)


def range_ref_precedents_may_compress(
    graph: DependencyGraph,
    range_ref,
    config: TacoBuildConfig,
) -> bool:
    """Return whether every cell in a precedent `RangeRef` may compress."""
    return all(precedent_may_compress(graph, k, config) for k in range_ref.cell_keys())


def is_internal_node(
    graph: DependencyGraph,
    key: NodeKey,
    config: TacoBuildConfig,
) -> bool:
    """Formula node that is not a target or declared input."""
    node = graph.get_node(key)
    if node is None or node.is_leaf or not node.formula:
        return False
    return not node.is_target and key not in config.exclude_input_keys


def compressed_edge_violates_codegen_boundaries(
    graph: DependencyGraph,
    edge: CompressedEdge,
    config: TacoBuildConfig,
) -> str | None:
    """Return a reason when a compressed edge breaks codegen boundary rules."""
    for key in edge.dependent.cell_keys():
        node = graph.get_node(key)
        if config.exclude_targets and node is not None and node.is_target:
            return f"dependent {key!r} is a target"
        if key in config.exclude_input_keys:
            return f"dependent {key!r} is an input/setter boundary cell"
    if config.internal_only:
        for key in edge.dependent.cell_keys():
            if not is_internal_node(graph, key, config):
                return f"dependent {key!r} is not internal-only compressible"
    return None


def assert_codegen_index_boundaries(
    graph: DependencyGraph,
    index: TacoIndex,
    config: TacoBuildConfig,
) -> None:
    """Assert no compressed edge places targets or boundary inputs on the dependent side."""
    for edge in index.compressed_edges:
        reason = compressed_edge_violates_codegen_boundaries(graph, edge, config)
        assert reason is None, (
            f"Codegen boundary violation on {edge.dependent} <- {edge.precedent}: {reason}"
        )
