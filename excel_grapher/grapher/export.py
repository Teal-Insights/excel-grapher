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
    LightweightVizNodeColumns,
    LightweightVizPayload,
    LightweightVizStats,
    write_lightweight_viz_data,
    write_web_viz_html,
)
from .node import Node, NodeKey, NodeView
from .subgraph import select_path_induced_subgraph


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
        return "UNKNOWN" if node.normalized_formula is not None else "EMPTY"
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


def to_networkx(
    graph: DependencyGraph | GraphReadView,
    *,
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
):
    """Convert a dependency graph to a NetworkX DiGraph.

    Accepts `DependencyGraph` or any graph-like object with node iteration,
    dependency lookup, and edge attributes (for example `ProjectionResult`).

    NetworkX is an optional dependency. If not installed, raises ImportError with a
    helpful message.
    """
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


def to_graphviz(
    graph: DependencyGraph,
    *,
    label_fn: Callable[[NodeKey, Node | NodeView], str] | None = None,
    highlight: set[NodeKey] | None = None,
    rankdir: str = "TB",
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
) -> str:
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


def to_mermaid(
    graph: DependencyGraph,
    *,
    label_fn: Callable[[NodeKey, Node | NodeView], str] | None = None,
    max_nodes: int = 100,
    include_formula_on_nodes: bool = True,
    max_formula_length: int | None = 120,
) -> str:
    validate_max_formula_length(max_formula_length)

    def safe_id(key: str) -> str:
        # Mermaid node IDs can't contain many punctuation characters; keep it simple.
        return (
            key.replace("!", "_")
            .replace(" ", "_")
            .replace("-", "_")
            .replace("'", "")
            .replace('"', "")
            .replace(".", "_")
        )

    def escape_mermaid_label(label: str) -> str:
        return label.replace("\\", "\\\\").replace('"', '\\"')

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
        label = escape_mermaid_label(str(label_raw).replace("\n", "<br>"))
        # Box for leaves, rounded for formulas.
        shape = f'["{label}"]' if node.is_leaf else f'("{label}")'
        lines.append(f"  {safe_id(key)}{shape}")

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
                lines.append(f"  {safe_id(key)} --> {safe_id(resolved)}")
            else:
                guard_label = escape_mermaid_label(str(guard))
                lines.append(f'  {safe_id(key)} -.->|"{guard_label}"| {safe_id(resolved)}')

    return "\n".join(lines)


__all__ = [
    "LightweightVizLocalEdges",
    "LightweightVizModule",
    "LightweightVizModuleEdge",
    "LightweightVizNodeColumns",
    "LightweightVizPayload",
    "LightweightVizStats",
    "select_path_induced_subgraph",
    "to_graphviz",
    "to_mermaid",
    "to_networkx",
    "write_web_viz_html",
    "write_lightweight_viz_data",
]
