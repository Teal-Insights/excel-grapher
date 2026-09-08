#!/usr/bin/env python3
"""Demonstrate cross-sheet TACO compression on ``cross_sheet_taco_patterns.xlsx``.

Builds a dependency graph from ``Report`` targets that reference ``Data`` inputs,
prints the full TACO index, compares it to a **codegen-boundary** index (targets
and declared input ranges stay at cell granularity), evaluates the targets with
``FormulaEvaluator``, and optionally plots the full index next to the cell-level
graph.

Run from the repo root::

    uv run python examples/micro_workbooks/demo_cross_sheet_taco_index.py
    uv run python examples/micro_workbooks/demo_cross_sheet_taco_index.py --output cross_sheet_compare.png
"""

from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.address_keys import CellKey, RangeKey, format_key, parse_node_key
from excel_grapher.grapher import (
    TacoBuildConfig,
    build_taco_index,
)
from excel_grapher.grapher.export import to_networkx
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.parser import expand_range
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


def _input_keys_from_ranges(specs: list[str]) -> frozenset[str]:
    keys: set[str] = set()
    for spec in specs:
        parsed = parse_node_key(spec)
        if isinstance(parsed, CellKey):
            keys.add(str(parsed))
            continue
        if not isinstance(parsed, RangeKey):
            raise TypeError(f"expected cell or range spec, got {type(parsed).__name__}: {spec}")
        for sheet, a1 in expand_range(
            sheet=parsed.sheet,
            start_col=parsed.min_col,
            start_row=parsed.min_row,
            end_col=parsed.max_col,
            end_row=parsed.max_row,
            max_cells=10_000,
        ):
            keys.add(format_key(sheet, a1))
    return frozenset(keys)


def _print_codegen_boundary(graph: DependencyGraph, full_index: TacoIndex) -> None:
    input_keys = _input_keys_from_ranges(INPUT_RANGES)
    codegen_index = build_taco_index(
        graph, TacoBuildConfig.for_codegen(graph, input_keys=input_keys)
    )

    print("=== Codegen boundary TACO ===")
    print(
        "Targets and declared inputs stay at cell granularity; "
        "only internal formula columns would compress (none in this workbook)."
    )
    print(f"  declared input cells: {len(input_keys)}")
    print(
        f"  full TACO compressed edges: {len(full_index.compressed_edges)} "
        f"-> codegen boundary: {len(codegen_index.compressed_edges)}"
    )
    _print_edge_summary("Codegen-boundary index:", codegen_index)

    print("=== FormulaEvaluator ===")
    with FormulaEvaluator(graph) as ev:
        results = ev.evaluate(list(graph.target_keys()))
    print(f"  evaluated targets: {len(results)}")
    for key in sorted(results):
        print(f"    {key}: {results[key]}")
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
        # TACO infers stride patterns from raw `$` markers, which
        # normalization strips, so the raw formulas must be kept.
        store_raw_formula=True,
    )
    full_index = build_taco_index(graph)

    print(f"workbook: {WORKBOOK.name}")
    print(f"targets: {', '.join(TARGETS)}")
    print()
    _print_edge_summary("=== Full TACO index (analysis default) ===", full_index)

    _print_codegen_boundary(graph, full_index)

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
