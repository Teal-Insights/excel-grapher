#!/usr/bin/env python3
"""Demonstrate ``build_taco_index`` on ``taco_patterns.xlsx``.

Builds a dependency graph, derives a TACO range-pattern compression index from
it, prints a short summary, and plots the cell-level graph next to the
compressed range-pattern graph (NetworkX + Matplotlib).

Run from the repo root::

    uv run python examples/micro_workbooks/demo_taco_index.py
    uv run python examples/micro_workbooks/demo_taco_index.py --output taco_compare.png

Requires dev dependencies (``networkx``, ``matplotlib``).
"""

from __future__ import annotations

import argparse
from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.export import to_networkx
from excel_grapher.grapher.range_compression import (
    TacoIndex,
    build_taco_index,
    materialize_precedents,
)
from excel_grapher.grapher.range_compression.types import RangeRef

WORKBOOK = Path(__file__).with_name("taco_patterns.xlsx")
TARGETS = [
    "Patterns!D3:D7",  # RR
    "Patterns!F3:F7",  # RF
    "Patterns!H3:H7",  # FR
    "Patterns!K3:K7",  # FF (+ RR lookup key)
    "Patterns!P3:P7",  # RR-Chain
]

_PATTERN_COLORS = {
    "RR": "#1f77b4",
    "RF": "#ff7f0e",
    "FR": "#2ca02c",
    "FF": "#d62728",
    "RR-Chain": "#9467bd",
    "Single": "#7f7f7f",
}


def _range_label(ref: RangeRef) -> str:
    if ref.min_col == ref.max_col and ref.min_row == ref.max_row:
        return f"{ref.sheet}!{ref.min_col}{ref.min_row}"
    return f"{ref.sheet}!{ref.min_col}{ref.min_row}:{ref.max_col}{ref.max_row}"


def _build_compressed_digraph(index: TacoIndex):
    import networkx as nx

    graph = nx.DiGraph()
    for edge in index.compressed_edges:
        dep_label = _range_label(edge.dependent)
        prec_label = _range_label(edge.precedent)
        graph.add_node(dep_label, node_kind="range")
        graph.add_node(prec_label, node_kind="range")
        graph.add_edge(
            dep_label,
            prec_label,
            pattern=str(edge.meta.kind),
            label=str(edge.meta.kind),
        )
    for single in index.single_edges:
        graph.add_node(single.dependent, node_kind="cell")
        graph.add_node(single.precedent, node_kind="cell")
        graph.add_edge(
            single.dependent,
            single.precedent,
            pattern="Single",
            label="Single",
        )
    return graph


def _workbook_layout(graph, keys):
    import networkx as nx

    positions: dict[str, tuple[float, float]] = {}
    for key in keys:
        node = graph.nodes[key]
        col = node.get("column")
        row = node.get("row")
        if col is None or row is None:
            continue
        col_i = ord(col[0]) - ord("A")
        positions[key] = (float(col_i), -float(row))
    if len(positions) == len(keys):
        return positions
    return nx.spring_layout(graph, seed=0)


def _draw_cell_graph(ax, graph) -> None:
    import networkx as nx

    pos = _workbook_layout(graph, list(graph.nodes))
    node_colors = ["#ffd966" if graph.nodes[n].get("is_leaf") else "#8ecae6" for n in graph.nodes]
    nx.draw_networkx_nodes(
        graph,
        pos,
        ax=ax,
        node_size=450,
        node_color=node_colors,
        edgecolors="#333333",
        linewidths=0.6,
    )
    nx.draw_networkx_edges(
        graph,
        pos,
        ax=ax,
        arrows=True,
        arrowsize=12,
        width=0.8,
        edge_color="#666666",
        alpha=0.55,
        connectionstyle="arc3,rad=0.05",
    )
    labels = {n: f"{graph.nodes[n]['column']}{graph.nodes[n]['row']}" for n in graph.nodes}
    nx.draw_networkx_labels(graph, pos, labels=labels, ax=ax, font_size=7)
    ax.set_title(f"Cell graph ({graph.number_of_nodes()} nodes, {graph.number_of_edges()} edges)")
    ax.axis("off")


def _draw_compressed_graph(ax, graph) -> None:
    import networkx as nx

    pos = nx.spring_layout(graph, seed=1, k=1.4)
    nx.draw_networkx_nodes(
        graph,
        pos,
        ax=ax,
        node_size=1200,
        node_color="#f4a261",
        edgecolors="#333333",
        linewidths=0.8,
    )
    for u, v, data in graph.edges(data=True):
        color = _PATTERN_COLORS.get(data.get("pattern", "Single"), "#7f7f7f")
        nx.draw_networkx_edges(
            graph,
            pos,
            edgelist=[(u, v)],
            ax=ax,
            arrows=True,
            arrowsize=14,
            width=2.0,
            edge_color=color,
            connectionstyle="arc3,rad=0.08",
        )
        nx.draw_networkx_edge_labels(
            graph,
            pos,
            edge_labels={(u, v): data.get("label", "")},
            ax=ax,
            font_size=7,
            label_pos=0.45,
        )
    nx.draw_networkx_labels(graph, pos, ax=ax, font_size=7)
    ax.set_title(f"TACO index ({graph.number_of_nodes()} ranges, {graph.number_of_edges()} edges)")
    ax.axis("off")


def plot_side_by_side(cell_graph, index: TacoIndex, *, output: Path | None) -> None:
    import matplotlib.pyplot as plt

    compressed = _build_compressed_digraph(index)
    fig, axes = plt.subplots(1, 2, figsize=(18, 9))
    _draw_cell_graph(axes[0], cell_graph)
    _draw_compressed_graph(axes[1], compressed)
    fig.suptitle(f"TACO compression — {WORKBOOK.name}", fontsize=14)
    fig.tight_layout()
    if output is not None:
        fig.savefig(output, dpi=150, bbox_inches="tight")
        print(f"wrote {output}")
    else:
        plt.show()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Save the side-by-side figure to this path instead of opening a window",
    )
    args = parser.parse_args()

    graph = create_dependency_graph(
        WORKBOOK,
        TARGETS,
        load_values=False,
    )
    index = build_taco_index(graph)

    cell_edges = sum(len(graph.get_dependencies(key)) for key in graph)
    compressed = len(index.compressed_edges)
    print(f"workbook: {WORKBOOK.name}")
    print(f"nodes: {len(graph)}")
    print(f"cell edges |E|: {cell_edges}")
    print(f"compressed edges: {compressed}")
    print()
    print("compressed patterns:")
    for edge in index.compressed_edges:
        print(f"  {edge.meta.kind}: {edge.dependent} <- {edge.precedent}")
    print()
    sample = "Patterns!F5"
    precs = index.find_precedents(sample)
    print(f"find_precedents({sample!r}) -> {precs}")
    print(f"materialized: {sorted(materialize_precedents(index, sample))}")
    print()

    cell_nx = to_networkx(graph, include_formula_on_nodes=False)
    plot_side_by_side(cell_nx, index, output=args.output)


if __name__ == "__main__":
    main()
