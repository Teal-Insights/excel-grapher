"""Batteries-included lightweight workbook graph visualization (core + default overlays)."""

from __future__ import annotations

import heapq
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Literal

if TYPE_CHECKING:
    from excel_grapher.exporter.web_viz_layout import WebVizLayoutSpec

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import normalize_key, parse_address
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.lightweight_viz import (
    WEBVIZ_LOUVAIN_DIRECTED_OVERLAY_ID,
    LightweightVizCore,
    LightweightVizCoreNodeColumns,
    LightweightVizLayoutInput,
    LightweightVizModuleEdge,
    LightweightVizOverlay,
    LightweightVizPayload,
    VizLimits,
    _build_int_adjacencies,
    _edge_list_filtered,
    assemble_lightweight_viz_payload,
    build_lightweight_viz_core,
    derive_partition_modules_table,
)
from excel_grapher.grapher.node import Node

# --- Graph algorithms (module inference; exporter-local) --------------------


def _dfs_postorder_finish(adj: list[list[int]], n: int) -> list[int]:
    visited = [False] * n
    order: list[int] = []
    for start in range(n):
        if visited[start]:
            continue
        stack: list[tuple[int, int]] = [(start, 0)]
        visited[start] = True
        while stack:
            v, ni = stack[-1]
            nbrs = adj[v]
            if ni < len(nbrs):
                w = nbrs[ni]
                stack[-1] = (v, ni + 1)
                if not visited[w]:
                    visited[w] = True
                    stack.append((w, 0))
            else:
                stack.pop()
                order.append(v)
    return order


def _assign_components_reverse(adj_rev: list[list[int]], order_rev: list[int], n: int) -> list[int]:
    comp = [-1] * n
    label = 0
    for start in order_rev:
        if comp[start] >= 0:
            continue
        stack = [start]
        comp[start] = label
        while stack:
            v = stack.pop()
            for w in adj_rev[v]:
                if comp[w] < 0:
                    comp[w] = label
                    stack.append(w)
        label += 1
    return comp


def iterative_kosaraju_scc(adj_out: list[list[int]], n: int) -> list[int]:
    order = _dfs_postorder_finish(adj_out, n)
    adj_rev = [[] for _ in range(n)]
    for u in range(n):
        for v in adj_out[u]:
            adj_rev[v].append(u)
    for row in adj_rev:
        row.sort()
    comp = _assign_components_reverse(adj_rev, list(reversed(order)), n)
    return comp


def _remap_components(comp: list[int]) -> tuple[list[int], int]:
    mapping: dict[int, int] = {}
    out = [0] * len(comp)
    nxt = 0
    for i, c in enumerate(comp):
        if c not in mapping:
            mapping[c] = nxt
            nxt += 1
        out[i] = mapping[c]
    return out, nxt


def build_condensation_edges(
    adj: list[list[int]], n: int, comp: list[int], n_comp: int
) -> list[list[int]]:
    edges: set[tuple[int, int]] = set()
    for u in range(n):
        cu = comp[u]
        for v in adj[u]:
            cv = comp[v]
            if cu != cv:
                edges.add((cu, cv))
    cond = [[] for _ in range(n_comp)]
    for a, b in sorted(edges):
        cond[a].append(b)
    return cond


def condensation_indegree(adj_cond: list[list[int]], n_comp: int) -> list[int]:
    indeg = [0] * n_comp
    for u in range(n_comp):
        for v in adj_cond[u]:
            indeg[v] += 1
    return indeg


def kahn_toposort(adj: list[list[int]], n: int) -> list[int] | None:
    indeg = condensation_indegree(adj, n)
    heap = [i for i in range(n) if indeg[i] == 0]
    heapq.heapify(heap)
    order: list[int] = []
    while heap:
        u = heapq.heappop(heap)
        order.append(u)
        for v in adj[u]:
            indeg[v] -= 1
            if indeg[v] == 0:
                heapq.heappush(heap, v)
    if len(order) != n:
        return None
    return order


def longest_path_ranks(adj_cond: list[list[int]], n_comp: int) -> list[int]:
    preds = [[] for _ in range(n_comp)]
    for u in range(n_comp):
        for v in adj_cond[u]:
            preds[v].append(u)
    for row in preds:
        row.sort()

    topo = kahn_toposort(adj_cond, n_comp)
    if topo is None:
        return [0] * n_comp

    rank = [0] * n_comp
    for v in topo:
        pr = preds[v]
        if not pr:
            rank[v] = 0
        else:
            rank[v] = max(rank[u] + 1 for u in pr)
    return rank


@dataclass(frozen=True, slots=True)
class ModuleAnalysisResult:
    """Shared context for core layout + module overlay (computed once per graph)."""

    scc_count: int
    module_of: tuple[int, ...]
    node_rank: tuple[int, ...]
    module_edges: tuple[LightweightVizModuleEdge, ...]


def _module_overlay_from_analysis(
    core,
    ma: ModuleAnalysisResult,
    *,
    display_name: str,
    overlay_id: str = WEBVIZ_LOUVAIN_DIRECTED_OVERLAY_ID,
) -> LightweightVizOverlay:
    modules = derive_partition_modules_table(core, ma.module_of, ma.node_rank)
    data = {
        "scc_count": ma.scc_count,
        "node_module_id": list(ma.module_of),
        "node_rank": list(ma.node_rank),
        "module_edges": [
            {
                "source_module_id": e.source_module_id,
                "target_module_id": e.target_module_id,
                "unconditional_weight": e.unconditional_weight,
                "guarded_weight": e.guarded_weight,
            }
            for e in ma.module_edges
        ],
        "modules": [
            {
                "id": m.id,
                "node_count": m.node_count,
                "rank_min": m.rank_min,
                "rank_max": m.rank_max,
                "centroid_x": m.centroid_x,
                "centroid_y": m.centroid_y,
                "density_mode": m.density_mode,
            }
            for m in modules
        ],
    }
    return LightweightVizOverlay(
        overlay_id=overlay_id,
        schema_version=1,
        kind="partition",
        data=data,
        display_name=display_name,
        supplemental_stats={"module_edge_count": len(ma.module_edges)},
    )


def _dependency_graph_from_networkx(nx_graph: Any) -> DependencyGraph:
    dep_graph = DependencyGraph()
    reserved_node_fields = {
        "sheet",
        "column",
        "row",
        "formula",
        "normalized_formula",
        "value",
        "is_leaf",
        "label",
        "value_type",
    }

    for raw_key, attrs in nx_graph.nodes(data=True):
        if not isinstance(raw_key, str):
            raise TypeError(
                "to_web_viz_payload expects string node ids in NodeKey format (e.g. 'Sheet!A1')"
            )
        key = normalize_key(raw_key)
        sheet, cell = parse_address(key)
        column, row = fastpyxl.utils.cell.coordinate_from_string(cell)
        formula = attrs.get("formula")
        normalized_formula = attrs.get("normalized_formula")
        if normalized_formula is None and isinstance(formula, str):
            normalized_formula = formula
        is_leaf = bool(attrs.get("is_leaf", nx_graph.out_degree(raw_key) == 0))
        metadata = {k: v for k, v in attrs.items() if k not in reserved_node_fields}
        dep_graph.add_node(
            Node(
                sheet=sheet,
                column=column,
                row=int(row),
                formula=formula if isinstance(formula, str) else None,
                normalized_formula=normalized_formula
                if isinstance(normalized_formula, str)
                else None,
                value=attrs.get("value"),
                is_leaf=is_leaf,
                metadata=metadata,
            )
        )

    for raw_from, raw_to, edge_attrs in nx_graph.edges(data=True):
        if not isinstance(raw_from, str) or not isinstance(raw_to, str):
            raise TypeError("to_web_viz_payload expects string edge endpoints in NodeKey format")
        edge_kwargs: dict[str, Any] = {}
        if "provenance" in edge_attrs:
            edge_kwargs["provenance"] = edge_attrs["provenance"]
        dep_graph.add_edge(
            normalize_key(raw_from),
            normalize_key(raw_to),
            guard=edge_attrs.get("guard"),
            **edge_kwargs,
        )

    return dep_graph


def _module_assignment_directed_louvain(
    nx_graph: Any,
    *,
    keys: list[str],
    include_guarded_edges: bool,
    seed: int,
    weight_attr: str | None,
) -> tuple[int, ...]:
    try:
        import networkx as nx  # type: ignore[import-not-found]
    except Exception as e:  # pragma: no cover
        raise ImportError("networkx is required for to_web_viz_payload()") from e

    work = nx.DiGraph()
    work.add_nodes_from(keys)
    allowed = set(keys)
    for raw_u, raw_v, attrs in nx_graph.edges(data=True):
        if not isinstance(raw_u, str) or not isinstance(raw_v, str):
            continue
        u = normalize_key(raw_u)
        v = normalize_key(raw_v)
        if u not in allowed or v not in allowed:
            continue
        if not include_guarded_edges and attrs.get("guard") is not None:
            continue
        if weight_attr is not None and weight_attr in attrs:
            work.add_edge(u, v, **{weight_attr: attrs[weight_attr]})
        else:
            work.add_edge(u, v)

    communities = nx.community.louvain_communities(work, seed=seed, weight=weight_attr)
    key_id = {k: i for i, k in enumerate(keys)}
    normalized = [sorted(comm, key=lambda key: key_id[key]) for comm in communities]
    normalized.sort(key=lambda comm: key_id[comm[0]] if comm else len(keys))

    module_of = [-1] * len(keys)
    for module_id, comm in enumerate(normalized):
        for key in comm:
            module_of[key_id[key]] = module_id
    for i, module_id in enumerate(module_of):
        if module_id < 0:
            module_of[i] = len(normalized)
            normalized.append([keys[i]])

    return tuple(module_of)


def _analyze_modules_directed_louvain_for_viz(
    dep_graph: DependencyGraph,
    nx_graph: Any,
    *,
    include_guarded_edges_for_partition: bool,
    seed: int,
    weight_attr: str | None,
) -> ModuleAnalysisResult:
    keys = sorted(dep_graph)
    n = len(keys)
    if n == 0:
        return ModuleAnalysisResult(
            scc_count=0,
            module_of=tuple(),
            node_rank=tuple(),
            module_edges=tuple(),
        )

    key_id = {k: i for i, k in enumerate(keys)}
    uncond, _all_adj = _build_int_adjacencies(dep_graph, keys, key_id)

    comp_raw = iterative_kosaraju_scc(uncond, n)
    comp, n_comp = _remap_components(comp_raw)
    adj_cond = build_condensation_edges(uncond, n, comp, n_comp)
    scc_rank = longest_path_ranks(adj_cond, n_comp)
    node_rank = tuple(scc_rank[comp[i]] for i in range(n))

    module_of = _module_assignment_directed_louvain(
        nx_graph,
        keys=keys,
        include_guarded_edges=include_guarded_edges_for_partition,
        seed=seed,
        weight_attr=weight_attr,
    )

    all_edges = _edge_list_filtered(dep_graph, keys, key_id, include_guarded=True)
    mod_edge_map: dict[tuple[int, int], list[int]] = {}
    for u, v, guarded in all_edges:
        mu, mv = module_of[u], module_of[v]
        if mu == mv:
            continue
        pair = mod_edge_map.setdefault((mu, mv), [0, 0])
        if guarded:
            pair[1] += 1
        else:
            pair[0] += 1

    module_edges = tuple(
        LightweightVizModuleEdge(
            source_module_id=a,
            target_module_id=b,
            unconditional_weight=weights[0],
            guarded_weight=weights[1],
        )
        for (a, b), weights in sorted(mod_edge_map.items())
    )
    return ModuleAnalysisResult(
        scc_count=n_comp,
        module_of=module_of,
        node_rank=node_rank,
        module_edges=module_edges,
    )


def _layout_graph_from_networkx(
    nx_graph: Any,
    *,
    keys: list[str],
    include_guarded_edges: bool,
    weight_attr: str | None,
):
    import networkx as nx  # type: ignore[import-not-found]

    work = nx.DiGraph()
    work.add_nodes_from(keys)
    allowed = set(keys)
    for raw_u, raw_v, attrs in nx_graph.edges(data=True):
        if not isinstance(raw_u, str) or not isinstance(raw_v, str):
            continue
        u = normalize_key(raw_u)
        v = normalize_key(raw_v)
        if u not in allowed or v not in allowed:
            continue
        if not include_guarded_edges and attrs.get("guard") is not None:
            continue
        if weight_attr is not None and weight_attr in attrs:
            work.add_edge(u, v, **{weight_attr: attrs[weight_attr]})
        else:
            work.add_edge(u, v)
    return work


def _graphviz_layout_positions(work, prog: str) -> dict[str, tuple[float, float]]:
    try:
        from networkx.drawing.nx_agraph import graphviz_layout  # type: ignore[import-not-found]

        raw_pos = graphviz_layout(work, prog=prog)
    except Exception:
        try:
            from networkx.drawing.nx_pydot import graphviz_layout  # type: ignore[import-not-found]

            raw_pos = graphviz_layout(work, prog=prog)
        except Exception as e:
            raise ImportError(
                "graphviz layout requires pygraphviz or pydot with graphviz installed"
            ) from e

    return {k: (float(v[0]), float(v[1])) for k, v in raw_pos.items()}


def _compute_networkx_layout_positions(
    work,
    *,
    keys: list[str],
    layout_mode: Literal["spring", "forceatlas2", "multipartite", "graphviz_dot", "graphviz_sfdp"],
    node_rank: tuple[int, ...],
    seed: int,
    weight_attr: str | None,
) -> dict[str, tuple[float, float]]:
    import networkx as nx  # type: ignore[import-not-found]

    if layout_mode == "spring":
        pos = nx.spring_layout(work, seed=seed, weight=weight_attr)
        return {k: (float(v[0]), float(v[1])) for k, v in pos.items()}
    if layout_mode == "forceatlas2":
        pos = nx.forceatlas2_layout(work, seed=seed, weight=weight_attr)
        return {k: (float(v[0]), float(v[1])) for k, v in pos.items()}
    if layout_mode == "multipartite":
        layered = work.copy()
        for i, key in enumerate(keys):
            layered.nodes[key]["subset"] = int(node_rank[i])
        pos = nx.multipartite_layout(layered, subset_key="subset", align="horizontal")
        return {k: (float(v[0]), float(v[1])) for k, v in pos.items()}
    if layout_mode == "graphviz_dot":
        return _graphviz_layout_positions(work, prog="dot")
    if layout_mode == "graphviz_sfdp":
        return _graphviz_layout_positions(work, prog="sfdp")
    raise ValueError(f"Unsupported web viz layout mode: {layout_mode}")


def _apply_networkx_layout_to_core(
    core: LightweightVizCore,
    *,
    keys: list[str],
    positions: dict[str, tuple[float, float]],
) -> LightweightVizCore:
    x_coords: list[float] = []
    y_coords: list[float] = []
    for key in keys:
        x, y = positions.get(key, (0.0, 0.0))
        x_coords.append(float(x))
        y_coords.append(float(y))

    nodes = core.nodes
    node_cols = LightweightVizCoreNodeColumns(
        sheet_index=nodes.sheet_index,
        row=nodes.row,
        column=nodes.column,
        is_leaf=nodes.is_leaf,
        formula=nodes.formula,
        in_degree=nodes.in_degree,
        out_degree=nodes.out_degree,
        rank=nodes.rank,
        x=tuple(x_coords),
        y=tuple(y_coords),
        bucket_density=nodes.bucket_density,
    )
    return LightweightVizCore(
        stats=core.stats,
        sheets=core.sheets,
        nodes=node_cols,
        local_edges=core.local_edges,
        max_local_nodes=core.max_local_nodes,
        max_local_edges=core.max_local_edges,
    )


WebVizPayload = LightweightVizPayload


def to_web_viz_payload(
    nx_graph: Any,
    *,
    max_local_nodes: int | None = None,
    max_local_edges: int | None = None,
    include_guarded_edges: bool = True,
    include_guarded_edges_for_partition: bool = False,
    layout: WebVizLayoutSpec = "stratified_multipartite",
    layout_config: dict[str, Any] | None = None,
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
    seed: int = 0,
    weight_attr: str | None = None,
    include_module_overlay: bool = True,
) -> WebVizPayload:
    """Build a web-visualization payload from a NetworkX DiGraph.

    Layout is selected by `layout` (registered web layout plugin id or a direct
    `WebVizLayoutPlugin` callable). The default `stratified_multipartite` uses SCC-condensation longest-path rank on the vertical axis and
    Louvain community ordering on the horizontal axis when `include_module_overlay` is true.
    Other built-in ids include `spring`, `forceatlas2`, `multipartite` (NetworkX
    `multipartite_layout`), `graphviz_dot`, and `graphviz_sfdp`.

    Set `include_module_overlay=False` to skip the partition overlay (single module color;
    overview still draws local graph edges in the viewer).
    """
    from excel_grapher.exporter.web_viz_layout import (
        WebVizLayoutContext,
        run_web_viz_layout,
    )

    dep_graph = _dependency_graph_from_networkx(nx_graph)
    keys = sorted(dep_graph)
    limits = VizLimits(max_local_nodes=max_local_nodes, max_local_edges=max_local_edges)
    ctx = WebVizLayoutContext(
        dep_graph=dep_graph,
        nx_graph=nx_graph,
        keys=keys,
        limits=limits,
        include_guarded_edges=include_guarded_edges,
        include_guarded_edges_for_partition=include_guarded_edges_for_partition,
        include_module_overlay=include_module_overlay,
        include_formula_on_nodes=include_formula_on_nodes,
        max_formula_length=max_formula_length,
        seed=seed,
        weight_attr=weight_attr,
    )
    lay = run_web_viz_layout(ctx, layout, layout_config)
    li: LightweightVizLayoutInput | None
    if include_module_overlay and lay.module_analysis is not None:
        m = lay.module_analysis
        li = LightweightVizLayoutInput(m.module_of, m.node_rank)
    else:
        li = None
    core_base = build_lightweight_viz_core(
        dep_graph,
        limits=limits,
        layout_input=li,
        layout_mode="bfs",
        include_guarded_edges=include_guarded_edges,
        include_formula_on_nodes=include_formula_on_nodes,
        max_formula_length=max_formula_length,
    )
    core = _apply_networkx_layout_to_core(core_base, keys=keys, positions=lay.positions)
    overlays: list[LightweightVizOverlay] = []
    if include_module_overlay and lay.module_analysis is not None:
        overlays.append(
            _module_overlay_from_analysis(
                core,
                lay.module_analysis,
                display_name="Directed Louvain modules",
                overlay_id=WEBVIZ_LOUVAIN_DIRECTED_OVERLAY_ID,
            )
        )
    return assemble_lightweight_viz_payload(
        core,
        overlays,
        annotations=lay.annotations,
        viewer_hints=lay.viewer_hints,
    )
