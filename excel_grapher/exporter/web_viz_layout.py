"""Web layout plugins for to_web_viz_payload: single entry point per layout id, optional annotations/hints."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any, Literal, Protocol, TypeAlias, runtime_checkable

from excel_grapher.grapher.lightweight_viz import VizGraph, VizLimits

LAYOUT_CLUSTERED_FORCE = "clustered_force"
LAYOUT_SPRING = "spring"
LAYOUT_FORCEATLAS2 = "forceatlas2"
LAYOUT_GRAPHVIZ_DOT = "graphviz_dot"
LAYOUT_GRAPHVIZ_SFDP = "graphviz_sfdp"

_NX_SUBMODES: tuple[WebVizNxSubmode, ...] = (
    "spring",
    "forceatlas2",
    "graphviz_dot",
    "graphviz_sfdp",
)

WebVizNxSubmode = Literal["spring", "forceatlas2", "graphviz_dot", "graphviz_sfdp"]


class _NxGraphSlot:
    """Hold a caller-supplied NetworkX graph or build one on first access."""

    __slots__ = ("_value", "_factory")

    def __init__(
        self,
        value: Any | None = None,
        *,
        factory: Callable[[], Any] | None = None,
    ) -> None:
        self._value = value
        self._factory = factory

    def peek(self) -> Any | None:
        """Return the graph if it already exists, without building one."""
        return self._value

    def get(self) -> Any:
        """Return the NetworkX graph, materializing it if needed."""
        if self._value is None:
            if self._factory is None:
                raise ImportError("networkx is required for this web viz layout")
            self._value = self._factory()
        return self._value


@dataclass(frozen=True, slots=True)
class WebVizLayoutContext:
    """Inputs shared by web layout plugins.

    `nx_graph` is lazy when the payload was built from a `GraphReadView`.
    Built-in layouts that can run from `dep_graph` must not read `nx_graph`.
    """

    dep_graph: VizGraph
    keys: list[str]
    limits: VizLimits
    include_guarded_edges: bool
    include_guarded_edges_for_partition: bool
    include_module_overlay: bool
    include_formula_on_nodes: bool
    max_formula_length: int | None
    seed: int
    weight_attr: str | None
    _nx: _NxGraphSlot = field(repr=False, compare=False)

    @property
    def nx_graph(self) -> Any:
        """NetworkX DiGraph for plugins that need it.

        When the caller passed a `DiGraph`, this is that object. When the caller
        passed a `GraphReadView`, the first access runs `to_networkx`.
        """
        return self._nx.get()

    @property
    def provided_nx_graph(self) -> Any | None:
        """Caller-supplied NetworkX graph, or `None` if input was a graph view."""
        return self._nx.peek()


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


WebVizLayoutSpec: TypeAlias = str | WebVizLayoutPlugin

_plugins: dict[str, WebVizLayoutPlugin] = {}


def register_web_viz_layout(
    layout_id: str, plugin: WebVizLayoutPlugin, *, replace: bool = False
) -> None:
    if layout_id in _plugins and not replace:
        raise ValueError(f"duplicate web viz layout_id: {layout_id!r}")
    _plugins[layout_id] = plugin


def list_web_viz_layouts() -> tuple[str, ...]:
    return tuple(sorted(_plugins))


def resolve_web_viz_layout(layout: WebVizLayoutSpec) -> WebVizLayoutPlugin:
    """Look up a registered layout ID or return a direct layout plugin."""
    if isinstance(layout, str):
        if layout not in _plugins:
            raise ValueError(
                f"Unknown web viz layout: {layout!r}. Known: {', '.join(sorted(_plugins))}"
            )
        return _plugins[layout]
    return layout


def run_web_viz_layout(
    ctx: WebVizLayoutContext, layout: WebVizLayoutSpec, layout_config: dict[str, Any] | None
) -> WebVizLayoutResult:
    """Execute a registered layout ID or direct layout plugin."""
    cfg = dict(layout_config or {})
    return resolve_web_viz_layout(layout)(ctx, cfg)


def unregister_web_viz_layout(layout_id: str) -> None:
    """Remove a registered layout plugin (for tests and notebook cleanup)."""
    _plugins.pop(layout_id, None)


def _clustered_force(ctx: WebVizLayoutContext, layout_config: dict[str, Any]) -> WebVizLayoutResult:
    """Shared multilevel clustered force layout; clusters are directed Louvain modules.

    `layout_config["rank_pull"]` selects the vertical pull towards input depth
    (`"none"`, `"between"` (default), or `"everywhere"`).
    """
    from excel_grapher.exporter import lightweight_viz as lv
    from excel_grapher.grapher import viz_layout
    from excel_grapher.grapher.lightweight_viz import _build_int_adjacencies

    rank_pull = layout_config.get("rank_pull", viz_layout.DEFAULT_RANK_PULL)
    if rank_pull not in viz_layout.RANK_PULL_MODES:
        raise ValueError(
            f"rank_pull must be one of {viz_layout.RANK_PULL_MODES}, got {rank_pull!r}"
        )
    ma = lv._analyze_modules_directed_louvain_for_viz(
        ctx.dep_graph,
        keys=ctx.keys,
        include_guarded_edges_for_partition=ctx.include_guarded_edges_for_partition,
        seed=ctx.seed,
        weight_attr=ctx.weight_attr,
        nx_graph=ctx.provided_nx_graph,
    )
    n = len(ctx.keys)
    key_id = {k: i for i, k in enumerate(ctx.keys)}
    uncond, all_adj = _build_int_adjacencies(ctx.dep_graph, ctx.keys, key_id)
    selected = all_adj if ctx.include_guarded_edges else uncond
    depths = viz_layout.input_depths(n, [(u, v) for u in range(n) for v in uncond[u]])
    pos = viz_layout.clustered_force_layout(
        n,
        [(u, v) for u in range(n) for v in selected[u]],
        ma.module_of,
        depths=depths,
        rank_pull=rank_pull,
        seed=ctx.seed,
    )
    positions = {key: (float(pos[i][0]), float(pos[i][1])) for i, key in enumerate(ctx.keys)}
    return WebVizLayoutResult(
        positions=positions,
        module_analysis=ma if ctx.include_module_overlay else None,
        annotations={
            "layout": LAYOUT_CLUSTERED_FORCE,
            "rank_pull": rank_pull,
            "tier": viz_layout.viz_tier(n),
        },
        viewer_hints={},
    )


def _nx_submode(
    submode: WebVizNxSubmode,
) -> WebVizLayoutPlugin:
    def _impl(ctx: WebVizLayoutContext, layout_config: dict[str, Any]) -> WebVizLayoutResult:
        del layout_config
        from excel_grapher.exporter import lightweight_viz as lv

        work = lv._layout_work_graph(
            ctx.dep_graph,
            keys=ctx.keys,
            include_guarded_edges=ctx.include_guarded_edges,
            weight_attr=ctx.weight_attr,
            nx_graph=ctx.provided_nx_graph,
        )
        ma: Any | None
        if ctx.include_module_overlay:
            ma = lv._analyze_modules_directed_louvain_for_viz(
                ctx.dep_graph,
                keys=ctx.keys,
                include_guarded_edges_for_partition=ctx.include_guarded_edges_for_partition,
                seed=ctx.seed,
                weight_attr=ctx.weight_attr,
                nx_graph=ctx.provided_nx_graph,
            )
        else:
            ma = None

        pos = lv._compute_networkx_layout_positions(
            work,
            layout_mode=submode,
            seed=ctx.seed,
            weight_attr=ctx.weight_attr,
        )
        return WebVizLayoutResult(
            positions=pos,
            module_analysis=ma,
            annotations={"layout": submode, "layout_group": "networkx_drawing"},
            viewer_hints={},
        )

    return _impl


def _register_builtin_plugins() -> None:
    register_web_viz_layout(LAYOUT_CLUSTERED_FORCE, _clustered_force)
    for sid in _NX_SUBMODES:
        register_web_viz_layout(sid, _nx_submode(sid))


_register_builtin_plugins()

__all__ = [
    "LAYOUT_CLUSTERED_FORCE",
    "LAYOUT_SPRING",
    "LAYOUT_FORCEATLAS2",
    "LAYOUT_GRAPHVIZ_DOT",
    "LAYOUT_GRAPHVIZ_SFDP",
    "WebVizLayoutContext",
    "WebVizLayoutPlugin",
    "WebVizLayoutResult",
    "WebVizLayoutSpec",
    "list_web_viz_layouts",
    "register_web_viz_layout",
    "resolve_web_viz_layout",
    "run_web_viz_layout",
    "unregister_web_viz_layout",
]
