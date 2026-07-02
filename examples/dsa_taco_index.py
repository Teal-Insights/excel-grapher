#!/usr/bin/env python3
"""Generate the TACO dependency graph for ``dsa_model.xlsx``.

Builds a dependency graph from the DSA output targets with ``taco_index=True``,
prints the compressed TACO index, and optionally plots it next to the cell-level
graph.

Run from the repo root (after ``uv run python examples/build_dsa.py``)::

    uv run python examples/dsa_taco_index.py
    uv run python examples/dsa_taco_index.py --output dsa_taco_compare.png
"""

from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.export import to_networkx
from excel_grapher.grapher.range_compression import TacoIndex

WORKBOOK = Path(__file__).with_name("dsa_model.xlsx")

# Model outputs: baseline debt path, stress flags, and the dashboard summary.
TARGETS = [
    "'Debt Dynamics'!B9:P9",
    "'Sustainability Indicators'!B21:P21",
    "Dashboard!B6:B13",
]


def _print_edge_summary(title: str, index: TacoIndex) -> None:
    print(title)
    print(f"  compressed edges: {len(index.compressed_edges)}")
    print(f"  single edges: {len(index.single_edges)}")
    for edge in index.compressed_edges:
        cross = " (cross-sheet)" if edge.precedent.sheet != edge.dependent.sheet else ""
        print(f"    {edge.meta.kind}: {edge.dependent} <- {edge.precedent}{cross}")
    print()


def _load_plot_module():
    demo_path = Path(__file__).parent / "micro_workbooks" / "demo_taco_index.py"
    spec = importlib.util.spec_from_file_location("demo_taco_index", demo_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {demo_path}")
    demo_mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(demo_mod)
    return demo_mod


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument(
        "--no-plot",
        action="store_true",
        help="Skip the side-by-side matplotlib figure",
    )
    args = parser.parse_args()

    if not WORKBOOK.exists():
        raise SystemExit(f"{WORKBOOK} not found — run `uv run python examples/build_dsa.py` first.")

    graph = create_dependency_graph(
        WORKBOOK,
        TARGETS,
        load_values=True,
        max_depth=100,
        taco_index=True,
    )
    full_index = graph.taco_index
    if full_index is None:
        raise RuntimeError("Expected graph.taco_index")

    print(f"workbook: {WORKBOOK.name}")
    print(f"targets: {', '.join(TARGETS)}")
    print(f"graph nodes: {len(graph)}")
    print()
    _print_edge_summary("=== Full TACO index (analysis default) ===", full_index)
    print(
        "Note: TACO groups formula runs down a column (consecutive rows). The DSA\n"
        "model fills formulas across year columns within a row, so compression is\n"
        "limited to any column-oriented runs (e.g. stacked scenario/indicator rows).\n"
    )

    if args.no_plot:
        return

    demo_mod = _load_plot_module()
    demo_mod.plot_side_by_side(
        to_networkx(graph, include_formula_on_nodes=False),
        full_index,
        output=args.output,
    )


if __name__ == "__main__":
    main()
