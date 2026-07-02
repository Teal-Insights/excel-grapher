#!/usr/bin/env python3
"""Demonstrate row-autofill TACO compression on ``taco_row_patterns.xlsx``.

Builds a dependency graph with TACO range-pattern compression (row autofill
demos), prints a short summary, and plots the cell-level graph next to the
compressed range-pattern graph (NetworkX + Matplotlib).

Column-autofill demos: ``demo_taco_index.py`` on ``taco_patterns.xlsx``.

Run from the repo root (after building the fixture)::

    uv run python examples/micro_workbooks/build_taco_row_patterns_workbook.py
    uv run python examples/micro_workbooks/demo_taco_row_index.py
    uv run python examples/micro_workbooks/demo_taco_row_index.py --output taco_row_compare.png

Requires dev dependencies (``networkx``, ``matplotlib``).
"""

from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.export import to_networkx
from excel_grapher.grapher.range_compression import (
    materialize_precedents,
)

WORKBOOK = Path(__file__).with_name("taco_row_patterns.xlsx")
TARGETS = [
    "PatternsRow!F9:J9",  # RR
    "PatternsRow!W9:AA9",  # RF
    "PatternsRow!AI9:AM9",  # FR
    "PatternsRow!BA9:BE9",  # FF
    "PatternsRow!AS9:AW9",  # RR-Chain
]


def _load_plot_module():
    demo_path = Path(__file__).with_name("demo_taco_index.py")
    spec = importlib.util.spec_from_file_location("demo_taco_index", demo_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {demo_path}")
    demo_mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(demo_mod)
    return demo_mod


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Save the side-by-side figure to this path instead of opening a window",
    )
    args = parser.parse_args()

    if not WORKBOOK.exists():
        raise SystemExit(
            f"{WORKBOOK} not found — run "
            "`uv run python examples/micro_workbooks/build_taco_row_patterns_workbook.py` first."
        )

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
    sample = "PatternsRow!G9"
    precs = index.find_precedents(sample)
    print(f"find_precedents({sample!r}) -> {precs}")
    print(f"materialized: {sorted(materialize_precedents(index, sample))}")
    print()

    demo_mod = _load_plot_module()
    cell_nx = to_networkx(graph, include_formula_on_nodes=False)
    demo_mod.plot_side_by_side(
        cell_nx,
        index,
        output=args.output,
        workbook_name=WORKBOOK.name,
    )


if __name__ == "__main__":
    main()
