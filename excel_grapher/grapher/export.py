from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
from typing import Any

from .formula_label import (
    display_formula,
    truncate_formula_display,
    validate_max_formula_length,
)
from .graph import DependencyGraph, GraphReadView
from .guard import GuardExpr
from .lightweight_viz import (
    LightweightVizLocalEdges,
    LightweightVizModule,
    LightweightVizModuleEdge,
    LightweightVizPayload,
    write_lightweight_viz_data,
    write_web_viz_html,
)
from .node import Node, NodeKey, NodeView
from .series_graph import SeriesGraph
from .sheet_graph import SheetGraph
from .subgraph import select_path_induced_subgraph, select_shortest_path_subgraph


def _dot_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")


def _guard_label(g: GuardExpr) -> str:
    return _dot_escape(str(g))


def _node_display_label(
    key: NodeKey,
    node: Node | NodeView,
    *,
    label_fn: Callable[[NodeKey, Node | NodeView], str] | None,
    include_formula_on_nodes: bool,
    max_formula_length: int | None,
) -> str:
    base = label_fn(key, node) if label_fn is not None else str(key)
    formula = display_formula(node)
    if not include_formula_on_nodes or not formula:
        return base
    shown = truncate_formula_display(formula, max_formula_length)
    return f"{base}\n{shown}"


def _networkx_value_type(node: Node | NodeView) -> str:
    value = node.value
    if value is None:
        return "UNKNOWN" if node.has_formula else "EMPTY"
    # bool must be checked before int, since bool subclasses int.
    if isinstance(value, bool):
        return "BOOLEAN"
    if isinstance(value, (int, float)):
        return "NUMBER"
    if isinstance(value, datetime):
        return "DATETIME"
    if isinstance(value, str):
        return "ERROR" if value.startswith("#") else "STRING"
    return "UNKNOWN"


def _quotient_node_id(node: object) -> str:
    """Return the GraphViz/Mermaid/NetworkX id for a quotient-graph node."""
    series_id = getattr(node, "series_id", None)
    if isinstance(series_id, str):
        return series_id
    name = getattr(node, "name", None)
    if isinstance(name, str):
        return name
    raise TypeError(f"unsupported quotient node type: {type(node)!r}")


def _sheet_graph_to_networkx(graph: SheetGraph):
    """Convert a sheet graph to a NetworkX DiGraph of worksheet names."""
    try:
        import networkx as nx
    except Exception as e:  # pragma: no cover
        raise ImportError("networkx is not installed; add it to use to_networkx()") from e

    G = nx.DiGraph()
    for node in graph.nodes:
        G.add_node(
            node.name,
            sheet=node.name,
            cell_count=node.cell_count,
            formula_count=node.formula_count,
            leaf_count=node.leaf_count,
            label=node.name,
        )
    for edge in graph.edges:
        G.add_edge(
            edge.source,
            edge.target,
            weight=edge.edge_count,
            edge_count=edge.edge_count,
            guarded_count=edge.guarded_count,
        )
    return G


def _series_graph_to_networkx(graph: SeriesGraph):
    """Convert a series graph to a NetworkX DiGraph of series ids."""
    try:
        import networkx as nx
    except Exception as e:  # pragma: no cover
        raise ImportError("networkx is not installed; add it to use to_networkx()") from e

    G = nx.DiGraph()
    for node in graph.nodes:
        G.add_node(
            node.series_id,
            series_id=node.series_id,
            direction=node.direction,
            layout=node.layout,
            sheet=node.sheet,
            cell_count=node.cell_count,
            formula_count=node.formula_count,
            leaf_count=node.leaf_count,
            label=node.series_id,
        )
    for edge in graph.edges:
        G.add_edge(
            edge.source,
            edge.target,
            weight=edge.edge_count,
            edge_count=edge.edge_count,
            guarded_count=edge.guarded_count,
        )
    return G


def to_networkx(
    graph: DependencyGraph | GraphReadView | SheetGraph | SeriesGraph,
    *,
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
):
    """Convert a dependency, sheet, or series graph to a NetworkX DiGraph.

    Accepts `DependencyGraph`, any graph-like object with node iteration,
    dependency lookup, and edge attributes (for example `ProjectionResult`),
    a `SheetGraph` from `to_sheet_graph`, or a `SeriesGraph` from
    `to_series_graph`. Statement graphs and series drilldowns are converted
    with `excel_grapher.exporter.statement_graph_to_networkx` and
    `SeriesDrilldown.to_networkx`.

    NetworkX is an optional dependency. If not installed, raises ImportError with a
    helpful message.
    """
    kind = type(graph).__name__
    if kind in {"StatementGraph", "SeriesDrilldown"}:
        raise TypeError(
            "use excel_grapher.exporter.statement_graph_to_networkx() "
            "or SeriesDrilldown.to_networkx() for statement-graph analysis; "
            "grapher.to_networkx() accepts cell, sheet, and series graphs"
        )
    if isinstance(graph, SheetGraph):
        return _sheet_graph_to_networkx(graph)
    if isinstance(graph, SeriesGraph):
        return _series_graph_to_networkx(graph)

    validate_max_formula_length(max_formula_length)

    try:
        import networkx as nx
    except Exception as e:  # pragma: no cover
        raise ImportError("networkx is not installed; add it to use to_networkx()") from e

    G = nx.DiGraph()

    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        attrs: dict[str, Any] = {
            "sheet": node.sheet,
            "column": node.column,
            "row": node.row,
            "formula": node.formula,
            "value": node.value,
            "value_type": _networkx_value_type(node),
            "is_leaf": node.is_leaf,
            "label": _node_display_label(
                key,
                node,
                label_fn=None,
                include_formula_on_nodes=include_formula_on_nodes,
                max_formula_length=max_formula_length,
            ),
        }
        attrs.update(node.metadata)
        G.add_node(key, **attrs)

    for key in graph:
        for dep in graph.get_dependencies(key):
            edge = graph.get_edge_attrs(key, dep)
            resolved = graph.resolve_endpoint(dep)
            if resolved is None:
                continue
            edge_kwargs: dict[str, Any] = {}
            if edge.guard is not None:
                edge_kwargs["guard"] = edge.guard
            if edge.provenance is not None:
                edge_kwargs["provenance"] = edge.provenance
            G.add_edge(key, resolved, **edge_kwargs)

    return G


def _quotient_graph_to_graphviz(
    graph: SheetGraph | SeriesGraph,
    *,
    rankdir: str,
    graph_name: str,
) -> str:
    """Render a sheet or series quotient as GraphViz DOT."""
    lines: list[str] = [f"digraph {graph_name} {{", f"  rankdir={_dot_escape(rankdir)};"]
    for node in graph.nodes:
        name = _dot_escape(_quotient_node_id(node))
        lines.append(f'  "{name}" [label="{name}"];')
    for edge in graph.edges:
        src = _dot_escape(edge.source)
        dst = _dot_escape(edge.target)
        label = _dot_escape(str(edge.edge_count))
        attrs = f'label="{label}"'
        if edge.edge_count > 0 and edge.guarded_count == edge.edge_count:
            attrs = f"style=dashed {attrs}"
        lines.append(f'  "{src}" -> "{dst}" [{attrs}];')
    lines.append("}")
    return "\n".join(lines)


def to_graphviz(
    graph: DependencyGraph | SheetGraph | SeriesGraph,
    *,
    label_fn: Callable[[NodeKey, Node | NodeView], str] | None = None,
    highlight: set[NodeKey] | None = None,
    rankdir: str = "TB",
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
) -> str:
    """Render a cell graph, `SheetGraph`, or `SeriesGraph` as GraphViz DOT.

    Formula-label options apply to cell graphs. Sheet and series graphs emit
    one node per worksheet or bound series and label each consolidated edge
    with its cell-edge count. Fully guarded quotient edges are dashed.
    """
    if isinstance(graph, SheetGraph):
        return _quotient_graph_to_graphviz(graph, rankdir=rankdir, graph_name="sheet_dependencies")
    if isinstance(graph, SeriesGraph):
        return _quotient_graph_to_graphviz(graph, rankdir=rankdir, graph_name="series_dependencies")

    validate_max_formula_length(max_formula_length)

    lines: list[str] = ["digraph dependencies {", f"  rankdir={_dot_escape(rankdir)};"]

    for key in graph.keys(order="workbook"):
        node = graph.get_node(key)
        if node is None:
            continue
        label_raw = _node_display_label(
            key,
            node,
            label_fn=label_fn,
            include_formula_on_nodes=include_formula_on_nodes,
            max_formula_length=max_formula_length,
        )
        label = _dot_escape(str(label_raw))
        shape = "box" if node.is_leaf else "ellipse"
        style = ""
        if highlight is not None and key in highlight:
            style = " style=filled fillcolor=yellow"
        lines.append(f'  "{_dot_escape(key)}" [label="{label}" shape={shape}{style}];')

    for key in graph.keys(order="workbook"):
        for dep in graph.get_dependencies(key):
            resolved = graph.resolve_endpoint(dep)
            if resolved is None or resolved not in graph:
                continue
            guard = graph.get_edge_guard(key, dep)
            if guard is None:
                lines.append(f'  "{_dot_escape(key)}" -> "{_dot_escape(resolved)}";')
            else:
                lines.append(
                    f'  "{_dot_escape(key)}" -> "{_dot_escape(resolved)}"'
                    f' [style=dashed label="{_guard_label(guard)}"];'
                )

    lines.append("}")
    return "\n".join(lines)


def _safe_mermaid_id(key: str) -> str:
    """Return a Mermaid node id with punctuation stripped."""
    return (
        key.replace("!", "_")
        .replace(" ", "_")
        .replace("-", "_")
        .replace("'", "")
        .replace('"', "")
        .replace(".", "_")
    )


def _escape_mermaid_label(label: str) -> str:
    return label.replace("\\", "\\\\").replace('"', '\\"')


def _quotient_graph_to_mermaid(graph: SheetGraph | SeriesGraph, *, max_nodes: int) -> str:
    """Render a sheet or series quotient as a Mermaid flowchart."""
    lines: list[str] = ["flowchart TD"]
    nodes = graph.nodes[: max_nodes if max_nodes > 0 else 0]
    node_ids = {_quotient_node_id(node) for node in nodes}
    for node in nodes:
        node_id = _quotient_node_id(node)
        label = _escape_mermaid_label(node_id)
        lines.append(f'  {_safe_mermaid_id(node_id)}["{label}"]')
    if len(graph.nodes) > len(nodes):
        lines.append(f"  truncated[[...{len(graph.nodes) - len(nodes)} more nodes]]")
    for edge in graph.edges:
        if edge.source not in node_ids or edge.target not in node_ids:
            continue
        src = _safe_mermaid_id(edge.source)
        dst = _safe_mermaid_id(edge.target)
        label = _escape_mermaid_label(str(edge.edge_count))
        if edge.edge_count > 0 and edge.guarded_count == edge.edge_count:
            lines.append(f'  {src} -.->|"{label}"| {dst}')
        else:
            lines.append(f'  {src} -->|"{label}"| {dst}')
    return "\n".join(lines)


def to_mermaid(
    graph: DependencyGraph | SheetGraph | SeriesGraph,
    *,
    label_fn: Callable[[NodeKey, Node | NodeView], str] | None = None,
    max_nodes: int = 100,
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
) -> str:
    """Render a cell graph, `SheetGraph`, or `SeriesGraph` as a Mermaid flowchart.

    Formula-label options apply to cell graphs. Sheet and series graphs emit
    one node per worksheet or bound series and label each consolidated edge
    with its cell-edge count. Fully guarded quotient edges use a dashed arrow.
    """
    if isinstance(graph, (SheetGraph, SeriesGraph)):
        return _quotient_graph_to_mermaid(graph, max_nodes=max_nodes)

    validate_max_formula_length(max_formula_length)

    lines: list[str] = ["flowchart TD"]

    keys = graph.keys(order="workbook")
    node_keys = keys[: max_nodes if max_nodes > 0 else 0]

    for key in node_keys:
        node = graph.get_node(key)
        if node is None:
            continue
        label_raw = _node_display_label(
            key,
            node,
            label_fn=label_fn,
            include_formula_on_nodes=include_formula_on_nodes,
            max_formula_length=max_formula_length,
        )
        # Mermaid flowchart labels use <br> for line breaks inside shapes.
        label = _escape_mermaid_label(str(label_raw).replace("\n", "<br>"))
        # Box for leaves, rounded for formulas.
        shape = f'["{label}"]' if node.is_leaf else f'("{label}")'
        lines.append(f"  {_safe_mermaid_id(key)}{shape}")

    if len(keys) > len(node_keys):
        lines.append(f"  truncated[[...{len(keys) - len(node_keys)} more nodes]]")

    node_set = set(node_keys)
    for key in node_keys:
        for dep in graph.get_dependencies(key):
            resolved = graph.resolve_endpoint(dep)
            if resolved is None or resolved not in node_set:
                continue
            guard = graph.get_edge_guard(key, dep)
            if guard is None:
                lines.append(f"  {_safe_mermaid_id(key)} --> {_safe_mermaid_id(resolved)}")
            else:
                guard_label = _escape_mermaid_label(str(guard))
                lines.append(
                    f'  {_safe_mermaid_id(key)} -.->|"{guard_label}"| {_safe_mermaid_id(resolved)}'
                )

    return "\n".join(lines)


__all__ = [
    "LightweightVizLocalEdges",
    "LightweightVizModule",
    "LightweightVizModuleEdge",
    "LightweightVizPayload",
    "select_path_induced_subgraph",
    "select_shortest_path_subgraph",
    "SheetGraph",
    "SeriesGraph",
    "to_graphviz",
    "to_mermaid",
    "to_networkx",
    "write_web_viz_html",
    "write_lightweight_viz_data",
]
