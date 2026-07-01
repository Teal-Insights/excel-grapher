"""Greedy column-adjacent grouping for TACO compression."""

from __future__ import annotations

from collections import defaultdict

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

from .boundaries import dependent_may_compress
from .config import TacoBuildConfig


def column_adjacent_groups(
    graph: DependencyGraph,
    *,
    min_len: int = 2,
    config: TacoBuildConfig | None = None,
) -> list[list[NodeKey]]:
    """Return groups of formula nodes sharing a column with consecutive rows.

    When ``config`` excludes targets or limits compression to internal nodes,
    long column runs are split so only compressible cells form a group.
    """
    cfg = config or TacoBuildConfig()
    by_column: dict[tuple[str, str], list[NodeKey]] = defaultdict(list)
    for key in graph:
        node = graph.get_node(key)
        if node is None or node.is_leaf or not node.formula:
            continue
        by_column[(node.sheet, node.column)].append(key)

    groups: list[list[NodeKey]] = []
    for keys in by_column.values():
        keys.sort(key=lambda k: _node_row(graph, k))
        run: list[NodeKey] = []
        prev_row: int | None = None
        for key in keys:
            node = graph.get_node(key)
            if node is None:
                continue
            if prev_row is not None and node.row != prev_row + 1:
                groups.extend(_flush_compressible_runs(graph, run, min_len=min_len, config=cfg))
                run = []
            run.append(key)
            prev_row = node.row
        groups.extend(_flush_compressible_runs(graph, run, min_len=min_len, config=cfg))
    return groups


def _uses_boundary_splits(config: TacoBuildConfig) -> bool:
    return config.exclude_targets or bool(config.exclude_input_keys) or config.internal_only


def _flush_compressible_runs(
    graph: DependencyGraph,
    keys: list[NodeKey],
    *,
    min_len: int,
    config: TacoBuildConfig,
) -> list[list[NodeKey]]:
    if not keys:
        return []
    if not _uses_boundary_splits(config):
        if len(keys) >= min_len:
            return [keys]
        return []

    out: list[list[NodeKey]] = []
    run: list[NodeKey] = []
    for key in keys:
        if dependent_may_compress(graph, key, config):
            run.append(key)
        else:
            if len(run) >= min_len:
                out.append(run)
            run = []
    if len(run) >= min_len:
        out.append(run)
    return out


def _node_row(graph: DependencyGraph, key: NodeKey) -> int:
    node = graph.get_node(key)
    return node.row if node is not None else 0
