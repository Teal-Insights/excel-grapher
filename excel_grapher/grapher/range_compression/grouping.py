"""Greedy column-adjacent grouping for TACO compression."""

from __future__ import annotations

from collections import defaultdict

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey


def column_adjacent_groups(graph: DependencyGraph, *, min_len: int = 2) -> list[list[NodeKey]]:
    """Return groups of formula nodes sharing a column with consecutive rows."""
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
                if len(run) >= min_len:
                    groups.append(run)
                run = []
            run.append(key)
            prev_row = node.row
        if len(run) >= min_len:
            groups.append(run)
    return groups


def _node_row(graph: DependencyGraph, key: NodeKey) -> int:
    node = graph.get_node(key)
    return node.row if node is not None else 0
