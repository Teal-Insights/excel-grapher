#!/usr/bin/env python3
"""Demonstrate the ``taco_index`` flag on ``taco_patterns.xlsx``.

Builds a dependency graph with an optional TACO range-pattern index attached
(see the TACO paper: https://arxiv.org/pdf/2302.05482).

Run from the repo root::

    uv run python examples/micro_workbooks/demo_taco_index.py
"""

from __future__ import annotations

from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import materialize_precedents

WORKBOOK = Path(__file__).with_name("taco_patterns.xlsx")
TARGETS = [
    "Patterns!D3:D7",  # RR
    "Patterns!F3:F7",  # RF
    "Patterns!H3:H7",  # FR
    "Patterns!K3:K7",  # FF (+ RR lookup key)
    "Patterns!P3:P7",  # RR-Chain
]


def main() -> None:
    graph = create_dependency_graph(
        WORKBOOK,
        TARGETS,
        load_values=False,
        taco_index=True,
    )
    index = graph.taco_index
    if index is None:
        raise RuntimeError("Expected graph.taco_index when taco_index=True")

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


if __name__ == "__main__":
    main()
