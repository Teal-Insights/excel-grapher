#!/usr/bin/env python3
"""Report cell-level edge counts vs TACO compressed index for a workbook."""

from __future__ import annotations

import argparse
import time
from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.range_compression import build_taco_index


def _count_cell_edges(graph) -> int:
    return sum(len(graph.get_dependencies(key)) for key in graph)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, help="Path to an .xlsx workbook")
    parser.add_argument(
        "targets",
        nargs="+",
        help="Sheet-qualified target cells or ranges (e.g. Data!D3:D1000)",
    )
    parser.add_argument(
        "--load-values",
        action="store_true",
        help="Load cached Excel values while building the graph",
    )
    args = parser.parse_args()

    t0 = time.perf_counter()
    graph = create_dependency_graph(
        args.workbook,
        args.targets,
        load_values=args.load_values,
    )
    graph_build_s = time.perf_counter() - t0

    t0 = time.perf_counter()
    index = build_taco_index(graph)
    index_build_s = time.perf_counter() - t0

    cell_edges = _count_cell_edges(graph)
    compressed_edges = len(index.compressed_edges)
    single_edges = len(index.single_edges)
    index_edges = compressed_edges + single_edges

    print(f"workbook: {args.workbook}")
    print(f"nodes: {len(graph)}")
    print(f"cell_edges |E|: {cell_edges}")
    print(f"compressed_edges: {compressed_edges}")
    print(f"single_edges: {single_edges}")
    print(f"index_edges |E_compressed|: {index_edges}")
    if cell_edges:
        print(f"compression_ratio: {cell_edges / index_edges:.2f}x")
    print(f"graph_build_s: {graph_build_s:.4f}")
    print(f"index_build_s: {index_build_s:.4f}")


if __name__ == "__main__":
    main()
