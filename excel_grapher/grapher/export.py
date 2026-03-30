from __future__ import annotations

from collections.abc import Callable
from typing import Any

from .graph import DependencyGraph
from .guard import GuardExpr
from .node import Node, NodeKey


def _dot_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")


def _guard_label(g: GuardExpr) -> str:
    return _dot_escape(str(g))


def to_networkx(graph: DependencyGraph):
    """
    Convert DependencyGraph to a NetworkX DiGraph.

    NetworkX is an optional dependency. If not installed, raises ImportError with a
    helpful message.
    """
    try:
        import networkx as nx  # type: ignore[import-not-found]
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
            "value_type": node.value_type.name,
            "is_leaf": node.is_leaf,
        }
        attrs.update(node.metadata)
        G.add_node(key, **attrs)

    for key in graph:
        for dep in graph.dependencies(key):
            attrs = graph.edge_attrs(key, dep)
            G.add_edge(key, dep, **attrs)

    return G


def to_graphviz(
    graph: DependencyGraph,
    *,
    label_fn: Callable[[NodeKey, Node], str] | None = None,
    highlight: set[NodeKey] | None = None,
    rankdir: str = "TB",
) -> str:
    lines: list[str] = ["digraph dependencies {", f"  rankdir={_dot_escape(rankdir)};"]

    for key in sorted(graph):
        node = graph.get_node(key)
        if node is None:
            continue
        label_raw = label_fn(key, node) if label_fn is not None else key
        label = _dot_escape(str(label_raw))
        shape = "box" if node.is_leaf else "ellipse"
        style = ""
        if highlight is not None and key in highlight:
            style = " style=filled fillcolor=yellow"
        lines.append(f'  "{_dot_escape(key)}" [label="{label}" shape={shape}{style}];')

    for key in sorted(graph):
        for dep in sorted(graph.dependencies(key)):
            guard = graph.edge_attrs(key, dep).get("guard")
            if guard is None:
                lines.append(f'  "{_dot_escape(key)}" -> "{_dot_escape(dep)}";')
            else:
                lines.append(
                    f'  "{_dot_escape(key)}" -> "{_dot_escape(dep)}"'
                    f' [style=dashed label="{_guard_label(guard)}"];'
                )

    lines.append("}")
    return "\n".join(lines)


def to_mermaid(
    graph: DependencyGraph,
    *,
    label_fn: Callable[[NodeKey, Node], str] | None = None,
    max_nodes: int = 100,
) -> str:
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

    lines: list[str] = ["flowchart TD"]

    keys = sorted(list(graph))
    node_keys = keys[: max_nodes if max_nodes > 0 else 0]

    for key in node_keys:
        node = graph.get_node(key)
        if node is None:
            continue
        label_raw = label_fn(key, node) if label_fn is not None else key
        label = str(label_raw).replace('"', '\\"')
        # Box for leaves, rounded for formulas.
        shape = f"[{label}]" if node.is_leaf else f"({label})"
        lines.append(f"  {safe_id(key)}{shape}")

    if len(keys) > len(node_keys):
        lines.append(f"  truncated[[...{len(keys) - len(node_keys)} more nodes]]")

    node_set = set(node_keys)
    for key in node_keys:
        for dep in sorted(graph.dependencies(key)):
            if dep not in node_set:
                continue
            guard = graph.edge_attrs(key, dep).get("guard")
            if guard is None:
                lines.append(f"  {safe_id(key)} --> {safe_id(dep)}")
            else:
                # Mermaid dashed edge. Text labels are optional; keep simple and stable.
                lines.append(f"  {safe_id(key)} -.-> {safe_id(dep)}")

    return "\n".join(lines)
