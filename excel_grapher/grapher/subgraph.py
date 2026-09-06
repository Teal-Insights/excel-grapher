from __future__ import annotations

from collections import deque
from collections.abc import Iterable, Sequence

from excel_grapher.core.address_keys import normalize_key

from .graph import DependencyGraph
from .node import NodeKey, copy_node


def select_shortest_path_subgraph(
    graph: DependencyGraph,
    *,
    source_key: NodeKey,
    target_key: NodeKey,
    directed: bool = True,
    max_path_length: int | None = None,
) -> DependencyGraph:
    """Return the induced subgraph over all hop-shortest paths between two nodes.

    Edge direction on the returned graph follows `DependencyGraph` semantics:
    `A -> B` means `A` depends on `B`. `directed` affects search only.

    Args:
        graph: Dependency graph to search.
        source_key: Path start node.
        target_key: Path end node.
        directed: If `True`, walk outgoing dependency edges only. If `False`,
            treat each edge as bidirectional for reachability.
        max_path_length: Optional hop-count ceiling. `0` allows only the
            trivial same-key path.

    Returns:
        Induced subgraph of every hop-shortest path. Endpoints are always
        included. Returned edges keep original direction, guards, and
        provenance.

    Raises:
        ValueError: If a key is missing, `max_path_length` is negative, no
            path exists under the requested directionality, or the shortest
            path is longer than `max_path_length`. A directed miss that still
            has an undirected path mentions `directed=False`.
    """
    sources = _normalize_existing_keys(graph, [source_key], arg_name="source_key")
    targets = _normalize_existing_keys(graph, [target_key], arg_name="target_key")
    source = sources[0]
    target = targets[0]
    _validate_limits(max_path_length=max_path_length, max_paths=None)

    path_nodes = _collect_shortest_path_nodes(
        graph,
        source=source,
        target=target,
        directed=directed,
        max_path_length=max_path_length,
    )
    return _induced_dependency_subgraph(graph, path_nodes)


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

    Edge direction follows `DependencyGraph` semantics: `A -> B` means
    `A` depends on `B`.
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
    return graph.keys(order="workbook", source=keys)


def _validate_limits(*, max_path_length: int | None, max_paths: int | None) -> None:
    if max_path_length is not None and max_path_length < 0:
        raise ValueError("max_path_length must be >= 0 when provided")
    if max_paths is not None and max_paths <= 0:
        raise ValueError("max_paths must be > 0 when provided")


def _resolved_neighbors(graph: DependencyGraph, key: NodeKey, *, directed: bool) -> list[NodeKey]:
    raw_neighbors: Iterable[NodeKey] = graph.get_dependencies(key)
    if not directed:
        raw_neighbors = (*raw_neighbors, *graph.get_dependents(key))

    neighbors: list[NodeKey] = []
    seen: set[NodeKey] = set()
    for raw in raw_neighbors:
        resolved = graph.resolve_endpoint(raw)
        if resolved is None or resolved in seen:
            continue
        seen.add(resolved)
        neighbors.append(resolved)
    return neighbors


def _bfs_shortest_parents(
    graph: DependencyGraph,
    source: NodeKey,
    *,
    directed: bool,
) -> tuple[dict[NodeKey, int], dict[NodeKey, list[NodeKey]]]:
    dist: dict[NodeKey, int] = {source: 0}
    parents: dict[NodeKey, list[NodeKey]] = {source: []}
    q: deque[NodeKey] = deque([source])
    while q:
        current = q.popleft()
        next_dist = dist[current] + 1
        for neighbor in _resolved_neighbors(graph, current, directed=directed):
            prior = dist.get(neighbor)
            if prior is None:
                dist[neighbor] = next_dist
                parents[neighbor] = [current]
                q.append(neighbor)
                continue
            if prior == next_dist:
                parents[neighbor].append(current)
    return dist, parents


def _nodes_on_shortest_paths(
    parents: dict[NodeKey, list[NodeKey]], target: NodeKey
) -> set[NodeKey]:
    keep: set[NodeKey] = set()
    stack = [target]
    while stack:
        node = stack.pop()
        if node in keep:
            continue
        keep.add(node)
        stack.extend(parents[node])
    return keep


def _collect_shortest_path_nodes(
    graph: DependencyGraph,
    *,
    source: NodeKey,
    target: NodeKey,
    directed: bool,
    max_path_length: int | None,
) -> set[NodeKey]:
    if source == target:
        return {source}

    dist, parents = _bfs_shortest_parents(graph, source, directed=directed)
    hop_length = dist.get(target)
    if hop_length is None:
        if directed:
            undirected_dist, _undirected_parents = _bfs_shortest_parents(
                graph, source, directed=False
            )
            if target in undirected_dist:
                raise ValueError(
                    f"no directed path from {source} to {target}; "
                    f"pass directed=False to search undirected paths"
                )
        raise ValueError(f"no path from {source} to {target}")

    if max_path_length is not None and hop_length > max_path_length:
        raise ValueError(
            f"max_path_length limit exceeded while collecting shortest paths: {max_path_length}"
        )
    return _nodes_on_shortest_paths(parents, target)


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

        for raw_dep in graph.get_dependencies(current):
            dep = graph.resolve_endpoint(raw_dep)
            if dep is None or dep in visited:
                continue

            next_depth = len(path)
            if max_path_length is not None and next_depth > max_path_length:
                if dep in can_reach_target:
                    raise ValueError(
                        f"max_path_length limit exceeded while collecting "
                        f"source->target paths: {max_path_length}"
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
        for dependent in graph.get_dependents(node):
            resolved = graph.resolve_endpoint(dependent) or dependent
            if resolved not in seen:
                q.append(resolved)
    return seen


def _induced_dependency_subgraph(
    graph: DependencyGraph, keep_keys: set[NodeKey]
) -> DependencyGraph:
    sub = DependencyGraph()
    if graph.sheet_order is not None:
        sub.sheet_order = list(graph.sheet_order)
    if graph.leaf_classification is not None:
        sub.leaf_classification = dict(graph.leaf_classification)

    for key in graph.keys(order="workbook", source=keep_keys):
        node = graph._get_internal_node(key)
        if node is None:
            continue
        sub.add_node(copy_node(node))

    for from_key in graph.keys(order="workbook", source=keep_keys):
        for to_key in graph.get_dependencies(from_key):
            resolved = graph.resolve_endpoint(to_key)
            if to_key not in keep_keys and (resolved is None or resolved not in keep_keys):
                continue
            attrs = graph.get_edge_attrs(from_key, to_key)
            edge_kwargs = {}
            if attrs.provenance is not None:
                edge_kwargs["provenance"] = attrs.provenance
            sub.add_edge(from_key, to_key, guard=attrs.guard, **edge_kwargs)
    return sub
