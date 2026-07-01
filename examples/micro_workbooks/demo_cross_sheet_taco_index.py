#!/usr/bin/env python3
"""Demonstrate cross-sheet TACO compression on ``cross_sheet_taco_patterns.xlsx``.

Builds a dependency graph from ``Report`` targets that reference ``Data`` inputs,
prints the full TACO index, compares it to a **codegen-boundary** index (targets
and inputs stay at cell granularity), runs ``CodeGenerator``, and optionally
plots the full index next to the cell-level graph.

Run from the repo root::

    uv run python examples/micro_workbooks/demo_cross_sheet_taco_index.py
    uv run python examples/micro_workbooks/demo_cross_sheet_taco_index.py --output cross_sheet_compare.png
"""

from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path
from typing import Any

from excel_grapher import create_dependency_graph
from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import (
    TacoBuildConfig,
    build_taco_index,
    input_keys_from_graph,
)
from excel_grapher.grapher.export import to_networkx
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.range_compression import TacoIndex

WORKBOOK = Path(__file__).with_name("cross_sheet_taco_patterns.xlsx")
TARGETS = ["Report!D3:D7", "Report!F3:F7", "Report!H3:H7", "Report!K3:K7"]
INPUT_RANGES = [
    "Data!B3:C7",
    "Data!E3:E11",
    "Data!G3:G7",
    "Data!M3:N7",
    "Report!J3:J7",
]


def _print_edge_summary(title: str, index: TacoIndex) -> None:
    print(title)
    print(f"  compressed edges: {len(index.compressed_edges)}")
    print(f"  single edges: {len(index.single_edges)}")
    for edge in index.compressed_edges:
        cross = " (cross-sheet)" if edge.precedent.sheet != edge.dependent.sheet else ""
        print(f"    {edge.meta.kind}: {edge.dependent} <- {edge.precedent}{cross}")
    print()


def _print_codegen_example(graph: DependencyGraph, full_index: TacoIndex) -> None:
    cell_targets = graph.target_keys()
    generator = CodeGenerator(graph)
    inputs, constants = generator.classify_leaf_nodes(
        cell_targets,
        input_ranges=INPUT_RANGES,
        attach_to_graph=True,
    )

    codegen_index = build_taco_index(graph, TacoBuildConfig.for_codegen(graph))

    print("=== Codegen boundary TACO ===")
    print(
        "Targets and declared inputs stay at cell granularity; "
        "only internal formula columns would compress (none in this workbook)."
    )
    print(f"  classified inputs: {len(inputs)}")
    print(f"  classified constants: {len(constants)}")
    print(f"  leaf inputs from graph: {len(input_keys_from_graph(graph))}")
    print(
        f"  full TACO compressed edges: {len(full_index.compressed_edges)} "
        f"-> codegen boundary: {len(codegen_index.compressed_edges)}"
    )
    _print_edge_summary("Codegen-boundary index:", codegen_index)

    print("=== CodeGenerator (cell-level export today) ===")
    print(
        "CodeGenerator does not consume the TACO index yet; "
        "it emits one function per formula cell (or vectorized target blocks)."
    )
    code = generator.generate(cell_targets)
    formula_defs = [
        line
        for line in code.splitlines()
        if line.startswith("def _formula_") or line.strip().startswith("'''Formula:")
    ]
    print(f"  generated lines: {len(code.splitlines())}")
    print(f"  formula function defs: {len(formula_defs) // 2}")
    print("  sample defs:")
    for line in formula_defs[:6]:
        print(f"    {line.strip()}")
    print()

    namespace: dict[str, Any] = {}
    exec(code, namespace)  # noqa: S102 — trusted local demo script
    results = namespace["compute_all"]()
    print("  compute_all() target blocks:")
    for key in sorted(results):
        value = results[key]
        preview = value.flatten().tolist() if hasattr(value, "flatten") else value
        print(f"    {key}: {preview}")
    print()


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
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument(
        "--no-plot",
        action="store_true",
        help="Skip the side-by-side matplotlib figure",
    )
    args = parser.parse_args()

    graph = create_dependency_graph(
        WORKBOOK,
        TARGETS,
        load_values=True,
    )
    full_index = build_taco_index(graph)

    print(f"workbook: {WORKBOOK.name}")
    print(f"targets: {', '.join(TARGETS)}")
    print()
    _print_edge_summary("=== Full TACO index (analysis default) ===", full_index)

    _print_codegen_example(graph, full_index)

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
