"""Web layout plugins for to_web_viz_payload: single entry point per layout id, optional annotations/hints."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal, Protocol, runtime_checkable

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.lightweight_viz import VizLimits

LAYOUT_STRATIFIED_MULTIPARTITE = "stratified_multipartite"
LAYOUT_SPRING = "spring"
LAYOUT_FORCEATLAS2 = "forceatlas2"
LAYOUT_MULTIPARTITE = "multipartite"
LAYOUT_GRAPHVIZ_DOT = "graphviz_dot"
LAYOUT_GRAPHVIZ_SFDP = "graphviz_sfdp"

_NX_SUBMODES: tuple[WebVizNxSubmode, ...] = (
    "spring",
    "forceatlas2",
    "multipartite",
    "graphviz_dot",
    "graphviz_sfdp",
)

WebVizNxSubmode = Literal["spring", "forceatlas2", "multipartite", "graphviz_dot", "graphviz_sfdp"]


@dataclass(frozen=True, slots=True)
class WebVizLayoutContext:
    dep_graph: DependencyGraph
    nx_graph: Any
    keys: list[str]
    limits: VizLimits
    include_guarded_edges: bool
    include_guarded_edges_for_partition: bool
    include_module_overlay: bool
    include_formula_on_nodes: bool
    max_formula_length: int | None
    seed: int
    weight_attr: str | None


@dataclass(frozen=True, slots=True)
class WebVizLayoutResult:
    """Output of a web layout plugin: node positions plus optional analysis for overlays."""

    positions: dict[str, tuple[float, float]]
    module_analysis: Any
    annotations: dict[str, Any]
    viewer_hints: dict[str, Any]


@runtime_checkable
class WebVizLayoutPlugin(Protocol):
    def __call__(
        self, ctx: WebVizLayoutContext, layout_config: dict[str, Any]
    ) -> WebVizLayoutResult: ...


_plugins: dict[str, WebVizLayoutPlugin] = {}


def register_web_viz_layout(layout_id: str, plugin: WebVizLayoutPlugin) -> None:
    if layout_id in _plugins:
        raise ValueError(f"duplicate web viz layout_id: {layout_id!r}")
    _plugins[layout_id] = plugin


def list_web_viz_layouts() -> tuple[str, ...]:
    return tuple(sorted(_plugins))


def run_web_viz_layout(
    ctx: WebVizLayoutContext, layout_id: str, layout_config: dict[str, Any] | None
) -> WebVizLayoutResult:
    if layout_id not in _plugins:
        raise ValueError(
            f"Unknown web viz layout: {layout_id!r}. Known: {', '.join(sorted(_plugins))}"
        )
    cfg = dict(layout_config or {})
    return _plugins[layout_id](ctx, cfg)


def _positions_stratified_scc_louvain(
    keys: list[str],
    ma: Any,
) -> dict[str, tuple[float, float]]:
    n = len(keys)
    if n == 0:
        return {}
    rank = ma.node_rank
    mod = ma.module_of
    max_r = max(rank) if n else 0
    by_rank: dict[int, list[int]] = {}
    for i in range(n):
        r = int(rank[i])
        by_rank.setdefault(r, []).append(i)
    out: dict[str, tuple[float, float]] = {}
    for r, idxs in sorted(by_rank.items(), key=lambda x: x[0]):
        idxs.sort(key=lambda i: (int(mod[i]), keys[i]))
        w = len(idxs)
        denom_y = max(max_r, 0) + 1
        y = 1.0 - 2.0 * (r + 0.5) / max(denom_y, 1)
        if w == 1:
            i0 = idxs[0]
            out[keys[i0]] = (0.0, y)
        else:
            for j, i in enumerate(idxs):
                x = 2.0 * j / (w - 1) - 1.0
                out[keys[i]] = (x, y)
    return out


def _stratified_multipartite(
    ctx: WebVizLayoutContext, layout_config: dict[str, Any]
) -> WebVizLayoutResult:
    del layout_config
    if not ctx.include_module_overlay:
        from excel_grapher.exporter import lightweight_viz as lv

        layout_graph = lv._layout_graph_from_networkx(
            ctx.nx_graph,
            keys=ctx.keys,
            include_guarded_edges=ctx.include_guarded_edges,
            weight_attr=ctx.weight_attr,
        )
        from excel_grapher.grapher.lightweight_viz import build_lightweight_viz_core

        core_tmp = build_lightweight_viz_core(
            ctx.dep_graph,
            limits=ctx.limits,
            layout_input=None,
            layout_mode="bfs",
            include_guarded_edges=ctx.include_guarded_edges,
            include_formula_on_nodes=ctx.include_formula_on_nodes,
            max_formula_length=ctx.max_formula_length,
        )
        node_rank = tuple(core_tmp.nodes.rank)
        pos = lv._compute_networkx_layout_positions(
            layout_graph,
            keys=ctx.keys,
            layout_mode="multipartite",
            node_rank=node_rank,
            seed=ctx.seed,
            weight_attr=ctx.weight_attr,
        )
        return WebVizLayoutResult(
            positions=pos,
            module_analysis=None,
            annotations={"layout": LAYOUT_MULTIPARTITE, "stratified_fallback": "bfs_multipartite"},
            viewer_hints={},
        )
    from excel_grapher.exporter import lightweight_viz as lv

    ma = lv._analyze_modules_directed_louvain_for_viz(
        ctx.dep_graph,
        ctx.nx_graph,
        include_guarded_edges_for_partition=ctx.include_guarded_edges_for_partition,
        seed=ctx.seed,
        weight_attr=ctx.weight_attr,
    )
    pos = _positions_stratified_scc_louvain(ctx.keys, ma)
    return WebVizLayoutResult(
        positions=pos,
        module_analysis=ma,
        annotations={"layout": LAYOUT_STRATIFIED_MULTIPARTITE},
        viewer_hints={},
    )


def _nx_submode(
    submode: WebVizNxSubmode,
) -> WebVizLayoutPlugin:
    def _impl(ctx: WebVizLayoutContext, layout_config: dict[str, Any]) -> WebVizLayoutResult:
        del layout_config
        from excel_grapher.exporter import lightweight_viz as lv
        from excel_grapher.grapher.lightweight_viz import build_lightweight_viz_core

        work = lv._layout_graph_from_networkx(
            ctx.nx_graph,
            keys=ctx.keys,
            include_guarded_edges=ctx.include_guarded_edges,
            weight_attr=ctx.weight_attr,
        )
        ma: Any | None
        if ctx.include_module_overlay:
            ma = lv._analyze_modules_directed_louvain_for_viz(
                ctx.dep_graph,
                ctx.nx_graph,
                include_guarded_edges_for_partition=ctx.include_guarded_edges_for_partition,
                seed=ctx.seed,
                weight_attr=ctx.weight_attr,
            )
            node_rank: tuple[int, ...] = ma.node_rank
        else:
            ma = None
            core_tmp = build_lightweight_viz_core(
                ctx.dep_graph,
                limits=ctx.limits,
                layout_input=None,
                layout_mode="bfs",
                include_guarded_edges=ctx.include_guarded_edges,
                include_formula_on_nodes=ctx.include_formula_on_nodes,
                max_formula_length=ctx.max_formula_length,
            )
            node_rank = tuple(core_tmp.nodes.rank)

        pos = lv._compute_networkx_layout_positions(
            work,
            keys=ctx.keys,
            layout_mode=submode,
            node_rank=node_rank,
            seed=ctx.seed,
            weight_attr=ctx.weight_attr,
        )
        if ctx.include_module_overlay:
            assert ma is not None
            return WebVizLayoutResult(
                positions=pos,
                module_analysis=ma,
                annotations={"layout": submode, "layout_group": "networkx_drawing"},
                viewer_hints={},
            )
        return WebVizLayoutResult(
            positions=pos,
            module_analysis=None,
            annotations={"layout": submode, "layout_group": "networkx_drawing"},
            viewer_hints={},
        )

    return _impl


def _register_builtin_plugins() -> None:
    register_web_viz_layout(LAYOUT_STRATIFIED_MULTIPARTITE, _stratified_multipartite)
    for sid in _NX_SUBMODES:
        register_web_viz_layout(sid, _nx_submode(sid))


_register_builtin_plugins()

__all__ = [
    "LAYOUT_STRATIFIED_MULTIPARTITE",
    "LAYOUT_SPRING",
    "LAYOUT_FORCEATLAS2",
    "LAYOUT_MULTIPARTITE",
    "LAYOUT_GRAPHVIZ_DOT",
    "LAYOUT_GRAPHVIZ_SFDP",
    "WebVizLayoutContext",
    "WebVizLayoutResult",
    "list_web_viz_layouts",
    "register_web_viz_layout",
    "run_web_viz_layout",
]
