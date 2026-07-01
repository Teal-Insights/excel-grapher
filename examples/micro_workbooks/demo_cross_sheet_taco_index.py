#!/usr/bin/env python3
"""Demonstrate cross-sheet TACO compression on ``cross_sheet_taco_patterns.xlsx``.

Run from the repo root::

    uv run python examples/micro_workbooks/demo_cross_sheet_taco_index.py
    uv run python examples/micro_workbooks/demo_cross_sheet_taco_index.py --output cross_sheet_compare.png
"""

from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.export import to_networkx

WORKBOOK = Path(__file__).with_name("cross_sheet_taco_patterns.xlsx")
TARGETS = ["Report!D3:D7", "Report!F3:F7", "Report!H3:H7", "Report!K3:K7"]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args()

    graph = create_dependency_graph(
        WORKBOOK,
        TARGETS,
        load_values=False,
        taco_index=True,
    )
    index = graph.taco_index
    if index is None:
        raise RuntimeError("Expected graph.taco_index")

    print(f"workbook: {WORKBOOK.name}")
    print(f"compressed edges: {len(index.compressed_edges)}")
    print()
    for edge in index.compressed_edges:
        cross = " (cross-sheet)" if edge.precedent.sheet != edge.dependent.sheet else ""
        print(f"  {edge.meta.kind}: {edge.dependent} <- {edge.precedent}{cross}")

    demo_path = Path(__file__).with_name("demo_taco_index.py")
    spec = importlib.util.spec_from_file_location("demo_taco_index", demo_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {demo_path}")
    demo_mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(demo_mod)

    demo_mod.plot_side_by_side(
        to_networkx(graph, include_formula_on_nodes=False),
        index,
        output=args.output,
    )


if __name__ == "__main__":
    main()
