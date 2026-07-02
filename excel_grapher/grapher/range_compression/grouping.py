"""Greedy column- and row-adjacent grouping for TACO compression."""

from __future__ import annotations

from collections import defaultdict
from enum import StrEnum

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

from .boundaries import dependent_may_compress
from .config import TacoBuildConfig


class Orientation(StrEnum):
    """Axis along which formula runs are grouped for compression."""

    column = "column"
    row = "row"


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
    return _adjacent_groups(
        graph,
        Orientation.column,
        min_len=min_len,
        config=config,
    )


def row_adjacent_groups(
    graph: DependencyGraph,
    *,
    min_len: int = 2,
    config: TacoBuildConfig | None = None,
) -> list[list[NodeKey]]:
    """Return groups of formula nodes sharing a row with consecutive columns.

    When ``config`` excludes targets or limits compression to internal nodes,
    long row runs are split so only compressible cells form a group.
    """
    return _adjacent_groups(
        graph,
        Orientation.row,
        min_len=min_len,
        config=config,
    )


def adjacent_groups(
    graph: DependencyGraph,
    *,
    min_len: int = 2,
    config: TacoBuildConfig | None = None,
    column_first: bool = True,
) -> list[list[NodeKey]]:
    """Return column- and row-adjacent groups without double-covering cells.

    When both orientations apply to the same cell, the primary orientation
    (column when ``column_first`` is True) claims it first.
    """
    cfg = config or TacoBuildConfig()
    if column_first:
        order = (Orientation.column, Orientation.row)
    else:
        order = (Orientation.row, Orientation.column)

    claimed: set[NodeKey] = set()
    groups: list[list[NodeKey]] = []
    for orientation in order:
        exclude = frozenset(claimed)
        for group in _adjacent_groups(
            graph,
            orientation,
            min_len=min_len,
            config=cfg,
            exclude_keys=exclude,
        ):
            groups.append(group)
            claimed.update(group)
    return groups


def _adjacent_groups(
    graph: DependencyGraph,
    orientation: Orientation,
    *,
    min_len: int = 2,
    config: TacoBuildConfig | None = None,
    exclude_keys: frozenset[NodeKey] = frozenset(),
) -> list[list[NodeKey]]:
    cfg = config or TacoBuildConfig()
    by_bucket: dict[tuple[str, str] | tuple[str, int], list[NodeKey]] = defaultdict(list)
    for key in graph:
        node = graph.get_node(key)
        if node is None or node.is_leaf or not node.formula:
            continue
        if orientation is Orientation.column:
            by_bucket[(node.sheet, node.column)].append(key)
        else:
            by_bucket[(node.sheet, node.row)].append(key)

    groups: list[list[NodeKey]] = []
    for keys in by_bucket.values():
        if orientation is Orientation.column:
            keys.sort(key=lambda k: _node_row(graph, k))
        else:
            keys.sort(key=lambda k: _node_col_index(graph, k))

        run: list[NodeKey] = []
        prev_axis: int | None = None
        for key in keys:
            if key in exclude_keys:
                if run:
                    groups.extend(_flush_compressible_runs(graph, run, min_len=min_len, config=cfg))
                    run = []
                prev_axis = None
                continue

            node = graph.get_node(key)
            if node is None:
                continue

            axis = node.row if orientation is Orientation.column else node.column_index
            if prev_axis is not None and axis != prev_axis + 1:
                groups.extend(_flush_compressible_runs(graph, run, min_len=min_len, config=cfg))
                run = []
            run.append(key)
            prev_axis = axis
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


def _node_col_index(graph: DependencyGraph, key: NodeKey) -> int:
    node = graph.get_node(key)
    return node.column_index if node is not None else 0
