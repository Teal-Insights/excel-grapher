from __future__ import annotations

from collections import deque
from collections.abc import Iterable, Sequence

from excel_grapher.core.address_keys import normalize_key

from .graph import DependencyGraph
from .node import Node, NodeKey


def select_path_induced_subgraph(
    graph: DependencyGraph,
    *,
    source_keys: Sequence[NodeKey],
    target_keys: Sequence[NodeKey],
    max_path_length: int | None = None,
    max_paths: int | None = None,
    include_endpoints: bool = True,
) -> DependencyGraph:
    """Return the induced subgraph over nodes lying on directed source->target paths.

    Edge direction follows ``DependencyGraph`` semantics: ``A -> B`` means
    ``A`` depends on ``B``.
    """

    sources = _normalize_existing_keys(graph, source_keys, arg_name="source_keys")
    targets = _normalize_existing_keys(graph, target_keys, arg_name="target_keys")
    _validate_limits(max_path_length=max_path_length, max_paths=max_paths)

    path_nodes = _collect_path_nodes(
        graph,
        sources=sources,
        targets=targets,
        max_path_length=max_path_length,
        max_paths=max_paths,
        include_endpoints=include_endpoints,
    )
    return _induced_dependency_subgraph(graph, path_nodes)


def _normalize_existing_keys(
    graph: DependencyGraph,
    raw_keys: Sequence[NodeKey],
    *,
    arg_name: str,
) -> list[NodeKey]:
    if not raw_keys:
        raise ValueError(f"{arg_name} cannot be empty")

    keys: list[NodeKey] = []
    seen: set[NodeKey] = set()
    missing: list[NodeKey] = []
    for raw in raw_keys:
        nk = normalize_key(raw)
        if nk in seen:
            continue
        seen.add(nk)
        if nk not in graph:
            missing.append(nk)
            continue
        keys.append(nk)
    if missing:
        names = ", ".join(sorted(missing))
        raise ValueError(f"{arg_name} contains keys not present in graph: {names}")
    return sorted(keys)


def _validate_limits(*, max_path_length: int | None, max_paths: int | None) -> None:
    if max_path_length is not None and max_path_length < 0:
        raise ValueError("max_path_length must be >= 0 when provided")
    if max_paths is not None and max_paths <= 0:
        raise ValueError("max_paths must be > 0 when provided")


def _collect_path_nodes(
    graph: DependencyGraph,
    *,
    sources: Sequence[NodeKey],
    targets: Sequence[NodeKey],
    max_path_length: int | None,
    max_paths: int | None,
    include_endpoints: bool,
) -> set[NodeKey]:
    target_set = set(targets)
    can_reach_target = _reverse_reachable_nodes(graph, target_set)
    path_nodes: set[NodeKey] = set()
    path_count = 0

    def add_path_nodes(path: list[NodeKey]) -> None:
        nonlocal path_count
        path_count += 1
        if max_paths is not None and path_count > max_paths:
            raise ValueError(
                f"max_paths limit exceeded while collecting source->target paths: {max_paths}"
            )
        if include_endpoints:
            path_nodes.update(path)
            return
        if len(path) > 2:
            path_nodes.update(path[1:-1])

    def dfs(current: NodeKey, path: list[NodeKey], visited: set[NodeKey]) -> None:
        if current in target_set:
            add_path_nodes(path)

        for dep in sorted(graph.get_dependencies(current)):
            if dep in visited:
                continue

            next_depth = len(path)
            if max_path_length is not None and next_depth > max_path_length:
                if dep in can_reach_target:
                    raise ValueError(
                        f"max_path_length limit exceeded while collecting source->target paths: {max_path_length}"
                    )
                continue

            visited.add(dep)
            path.append(dep)
            dfs(dep, path, visited)
            path.pop()
            visited.remove(dep)

    for source in sources:
        dfs(source, [source], {source})
    return path_nodes


def _reverse_reachable_nodes(graph: DependencyGraph, seeds: Iterable[NodeKey]) -> set[NodeKey]:
    seen: set[NodeKey] = set()
    q = deque(seeds)
    while q:
        node = q.popleft()
        if node in seen:
            continue
        seen.add(node)
        for dep in graph.get_dependents(node):
            if dep not in seen:
                q.append(dep)
    return seen


def _induced_dependency_subgraph(
    graph: DependencyGraph, keep_keys: set[NodeKey]
) -> DependencyGraph:
    sub = DependencyGraph()
    if graph.leaf_classification is not None:
        sub.leaf_classification = dict(graph.leaf_classification)

    for key in sorted(keep_keys):
        node = graph._get_internal_node(key)
        if node is None:
            continue
        sub.add_node(
            Node(
                sheet=node.sheet,
                column=node.column,
                row=node.row,
                formula=node.formula,
                normalized_formula=node.normalized_formula,
                value=node.value,
                is_leaf=node.is_leaf,
                metadata=dict(node.metadata),
            )
        )

    for from_key in sorted(keep_keys):
        for to_key in sorted(graph.get_dependencies(from_key)):
            if to_key not in keep_keys:
                continue
            attrs = graph.get_edge_attrs(from_key, to_key)
            edge_kwargs = {}
            if attrs.provenance is not None:
                edge_kwargs["provenance"] = attrs.provenance
            sub.add_edge(from_key, to_key, guard=attrs.guard, **edge_kwargs)
    return sub
