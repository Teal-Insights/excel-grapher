#!/usr/bin/env python3
"""Compare `DependencyGraph` dict adjacency with a CSR+CSC sidecar (#908).

Does not switch production storage. Builds `excel_grapher.grapher.csr_adjacency`
from an existing graph, then reports exclusive bytes and representative walk
times.

Usage:
    uv run python scripts/measure_csr_adjacency.py
    uv run python scripts/measure_csr_adjacency.py --preset if-heavy
    uv run python scripts/measure_csr_adjacency.py --preset lic-dsf
    uv run python scripts/measure_csr_adjacency.py --preset synthetic --synthetic-nodes 4000
    uv run python scripts/measure_csr_adjacency.py --workbook book.xlsx --targets 'Sheet1!A1'
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import tracemalloc
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_grapher.grapher.csr_adjacency import (
    LANDING_API_BREAKS,
    CsrCscAdjacency,
    from_graph,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node
from scripts.measure_graph_memory import (
    DEFAULT_TARGETS,
    DEFAULT_WORKBOOK,
    IF_HEAVY_TARGETS,
    IF_HEAVY_WORKBOOK,
    _walk,
    measure_graph_memory,
)

_REPO_ROOT = Path(__file__).resolve().parents[1]
LIC_DSF_WORKBOOK = (
    _REPO_ROOT / "tests" / "fixtures" / "lic_dsf" / "lic-dsf-template-2025-08-12.xlsm"
)

_DESCRIPTION = "Compare DependencyGraph dict adjacency with a CSR+CSC sidecar."


@dataclass(frozen=True, slots=True)
class CsrMemoryBreakdown:
    """Exclusive-vs-shared split of the sidecar against the live graph."""

    array_bytes: int
    index_table_exclusive_bytes: int
    exclusive_bytes: int
    shared_bytes: int
    total_bytes: int
    object_count: int


@dataclass(frozen=True, slots=True)
class WalkTiming:
    """One timed walk over the same arc set."""

    name: str
    seconds: float
    repeats: int
    checksum: int


def _array_bytes(csr: CsrCscAdjacency) -> int:
    return sum(sys.getsizeof(part) for part in (csr.row_ptr, csr.col_idx, csr.col_ptr, csr.row_idx))


def measure_csr_memory(graph: DependencyGraph, csr: CsrCscAdjacency) -> CsrMemoryBreakdown:
    """Walk the sidecar and split bytes that are unique vs shared with `graph`."""
    graph_ids = set(_walk((graph._nodes, graph._edges, graph._reverse_edges, graph._guards)).keys())
    csr_walk = _walk((csr.keys, csr.index, csr.row_ptr, csr.col_idx, csr.col_ptr, csr.row_idx))
    exclusive = shared = count = 0
    for obj_id, info in csr_walk.items():
        if info.is_singleton:
            continue
        count += 1
        if obj_id in graph_ids:
            shared += info.size
        else:
            exclusive += info.size
    arrays = _array_bytes(csr)
    array_ids = {id(part) for part in (csr.row_ptr, csr.col_idx, csr.col_ptr, csr.row_idx)}
    index_exclusive = sum(
        info.size
        for obj_id, info in csr_walk.items()
        if not info.is_singleton and obj_id not in graph_ids and obj_id not in array_ids
    )
    return CsrMemoryBreakdown(
        array_bytes=arrays,
        index_table_exclusive_bytes=index_exclusive,
        exclusive_bytes=exclusive,
        shared_bytes=shared,
        total_bytes=exclusive + shared,
        object_count=count,
    )


def _repeat_until(fn, *, min_repeats: int, min_seconds: float) -> tuple[float, int, int]:
    start = time.perf_counter()
    checksum = fn()
    repeats = 1
    while repeats < min_repeats or (time.perf_counter() - start) < min_seconds:
        checksum = fn()
        repeats += 1
    elapsed = time.perf_counter() - start
    return elapsed, repeats, checksum


def _walk_dict_forward(graph: DependencyGraph) -> int:
    total = 0
    nodes = graph._nodes
    for _key, neighbors in graph._edges.items():
        for dep in neighbors:
            if dep in nodes:
                total += 1
    return total


def _walk_dict_reverse(graph: DependencyGraph) -> int:
    total = 0
    nodes = graph._nodes
    for key, neighbors in graph._reverse_edges.items():
        if key not in nodes:
            continue
        for dep in neighbors:
            if dep in nodes:
                total += 1
    return total


def _walk_get_dependencies(graph: DependencyGraph) -> int:
    return sum(len(graph.get_dependencies(key)) for key in graph)


def _walk_csr_api(csr: CsrCscAdjacency) -> int:
    return sum(len(csr.dependencies(key)) for key in csr.keys)


def _walk_eval_order(graph: DependencyGraph) -> int:
    return len(graph.evaluation_order(strict=False))


def _topo_csr(csr: CsrCscAdjacency) -> int:
    """DFS over CSR (no cycle_report); checksum is the visit count."""
    n = csr.n
    state = bytearray(n)  # 0 unseen, 1 temp, 2 perm
    ptr = csr.row_ptr
    col = csr.col_idx
    visits = 0

    def visit(i: int) -> None:
        nonlocal visits
        marker = state[i]
        if marker == 2:
            return
        if marker == 1:
            return
        state[i] = 1
        for k in range(ptr[i], ptr[i + 1]):
            visit(col[k])
        state[i] = 2
        visits += 1

    for i in range(n):
        if state[i] == 0:
            visit(i)
    return visits


def time_walks(
    graph: DependencyGraph,
    csr: CsrCscAdjacency,
    *,
    include_eval_order: bool,
    min_repeats: int = 3,
    min_seconds: float = 0.05,
) -> list[WalkTiming]:
    """Time dict vs CSR neighbor walks (and optional `evaluation_order`)."""
    specs: list[tuple[str, Any]] = [
        ("dict_forward_sets", lambda: _walk_dict_forward(graph)),
        ("csr_forward_ids", csr.visit_forward_ids),
        ("dict_reverse_sets", lambda: _walk_dict_reverse(graph)),
        ("csr_reverse_ids", csr.visit_reverse_ids),
        ("get_dependencies_frozenset", lambda: _walk_get_dependencies(graph)),
        ("csr_dependencies_frozenset", lambda: _walk_csr_api(csr)),
        ("csr_topo_dfs", lambda: _topo_csr(csr)),
    ]
    if include_eval_order:
        specs.append(("evaluation_order", lambda: _walk_eval_order(graph)))
    out: list[WalkTiming] = []
    for name, fn in specs:
        seconds, repeats, checksum = _repeat_until(
            fn, min_repeats=min_repeats, min_seconds=min_seconds
        )
        out.append(WalkTiming(name=name, seconds=seconds, repeats=repeats, checksum=checksum))
    return out


def _synthetic_graph(n_pairs: int, extra_degree: int) -> DependencyGraph:
    graph = DependencyGraph()
    for row in range(1, n_pairs + 1):
        graph.add_node(make_cell_node("Sheet1", "A", row, value=float(row), is_leaf=True))
        graph.add_node(make_cell_node("Sheet1", "B", row, normalized_formula="=1", is_leaf=False))
        graph.add_edge(f"Sheet1!B{row}", f"Sheet1!A{row}")
        for back in range(1, extra_degree + 1):
            src = row - back
            if src >= 1:
                graph.add_edge(f"Sheet1!B{row}", f"Sheet1!B{src}")
    return graph


def _build_graph(args: argparse.Namespace) -> tuple[str, DependencyGraph]:
    if args.preset == "synthetic" and args.workbook is None:
        graph = _synthetic_graph(args.synthetic_nodes, args.synthetic_degree)
        return (
            f"synthetic ({args.synthetic_nodes} pairs, degree {args.synthetic_degree})",
            graph,
        )

    from excel_grapher import create_dependency_graph

    max_depth = 50
    use_cached_dynamic_refs = bool(args.use_cached_dynamic_refs)

    if args.preset == "taco":
        workbook, targets = DEFAULT_WORKBOOK, list(DEFAULT_TARGETS)
    elif args.preset == "if-heavy":
        workbook, targets = IF_HEAVY_WORKBOOK, list(IF_HEAVY_TARGETS)
    elif args.preset == "lic-dsf":
        from tests.integration.evaluator.utils.lic_dsf_chart_targets import (
            GRAPH_MAX_DEPTH,
            collect_chart_data_cell_keys,
        )

        workbook = LIC_DSF_WORKBOOK
        targets = collect_chart_data_cell_keys()
        use_cached_dynamic_refs = True
        max_depth = GRAPH_MAX_DEPTH
    else:
        workbook = args.workbook if args.workbook is not None else DEFAULT_WORKBOOK
        targets = args.targets if args.targets else list(DEFAULT_TARGETS)

    if not workbook.is_file():
        raise FileNotFoundError(f"workbook not found: {workbook}")

    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=args.load_values,
        capture_dependency_provenance=not args.no_provenance,
        use_cached_dynamic_refs=use_cached_dynamic_refs,
        max_depth=max_depth,
    )
    return str(workbook), graph


def _human(value: int) -> str:
    if value >= 1024 * 1024:
        return f"{value / (1024 * 1024):,.2f} MiB"
    return f"{value / 1024:,.1f} KiB"


def _ratio(num: int, den: int) -> str:
    if den == 0:
        return "n/a"
    return f"{den / num:,.1f}x" if num else "n/a"


def render_report(
    *,
    source: str,
    graph: DependencyGraph,
    csr: CsrCscAdjacency,
    build_peak: int,
    walks: list[WalkTiming],
    include_eval_order: bool,
) -> str:
    report = measure_graph_memory(graph)
    fwd = report.component("edges_forward")
    rev = report.component("edges_reverse")
    adj_exclusive = fwd.exclusive_bytes + rev.exclusive_bytes
    adj_total = fwd.total_bytes + rev.total_bytes
    csr_mem = measure_csr_memory(graph, csr)
    lines = [
        f"source: {source}",
        f"DependencyGraph: {report.node_count:,} nodes, {report.edge_count:,} edges",
        f"CSR/CSC:         {csr.n:,} rows, {csr.nnz:,} in-graph arcs"
        f" (dropped {csr.dropped_endpoints:,} dangling endpoints)",
        "",
        "exclusive bytes (identity-keyed walk; NodeKey strings shared with `_nodes` are not"
        " charged to CSR):",
        f"  dict `_edges` exclusive              {fwd.exclusive_bytes:>14,} B"
        f" ({_human(fwd.exclusive_bytes)})",
        f"  dict `_reverse_edges` exclusive      {rev.exclusive_bytes:>14,} B"
        f" ({_human(rev.exclusive_bytes)})",
        f"  dict adjacency exclusive (sum)       {adj_exclusive:>14,} B ({_human(adj_exclusive)})",
        f"  dict adjacency component total       {adj_total:>14,} B ({_human(adj_total)})",
        f"  CSR+CSC arrays (row_ptr/col_idx + CSC){csr_mem.array_bytes:>13,} B"
        f" ({_human(csr_mem.array_bytes)})",
        f"  node-index table exclusive           {csr_mem.index_table_exclusive_bytes:>14,} B"
        f" ({_human(csr_mem.index_table_exclusive_bytes)})",
        f"  CSR sidecar exclusive (arrays+index) {csr_mem.exclusive_bytes:>14,} B"
        f" ({_human(csr_mem.exclusive_bytes)})",
        f"  ratio dict exclusive / CSR exclusive {_ratio(csr_mem.exclusive_bytes, adj_exclusive):>14}",
        "",
        "build:",
        f"  second Python edge index             {csr.stats.built_second_python_edge_index}",
        f"  held input `_edges` during build     {csr.stats.held_input_adjacency}",
        f"  uint32 scratch (O(n), not O(nnz))    {csr.stats.scratch_uint32:,} slots",
        f"  tracemalloc peak inside from_graph   {build_peak:>14,} B ({_human(build_peak)})",
        "",
        "walks (lower seconds/repeat is faster):",
    ]
    for walk in walks:
        per = walk.seconds / walk.repeats
        lines.append(
            f"  {walk.name:<32}{per:>12.6f} s/repeat  "
            f"({walk.repeats} repeats, checksum={walk.checksum})"
        )
    if not include_eval_order:
        lines.append(
            "  evaluation_order skipped (cycle_report rebuilds dict adjacency; "
            "pass --eval-order to force)."
        )
    lines += ["", "landing API breaks if this ships:"]
    lines += [f"  - {item}" for item in LANDING_API_BREAKS]
    return "\n".join(lines)


def report_to_dict(
    *,
    source: str,
    graph: DependencyGraph,
    csr: CsrCscAdjacency,
    build_peak: int,
    walks: list[WalkTiming],
) -> dict[str, Any]:
    report = measure_graph_memory(graph)
    fwd = report.component("edges_forward")
    rev = report.component("edges_reverse")
    csr_mem = measure_csr_memory(graph, csr)
    return {
        "source": source,
        "node_count": report.node_count,
        "edge_count": report.edge_count,
        "csr_n": csr.n,
        "csr_nnz": csr.nnz,
        "dropped_endpoints": csr.dropped_endpoints,
        "edges_forward_exclusive": fwd.exclusive_bytes,
        "edges_reverse_exclusive": rev.exclusive_bytes,
        "adjacency_exclusive": fwd.exclusive_bytes + rev.exclusive_bytes,
        "csr_array_bytes": csr_mem.array_bytes,
        "csr_index_table_exclusive_bytes": csr_mem.index_table_exclusive_bytes,
        "csr_exclusive_bytes": csr_mem.exclusive_bytes,
        "build_peak_bytes": build_peak,
        "built_second_python_edge_index": csr.stats.built_second_python_edge_index,
        "held_input_adjacency": csr.stats.held_input_adjacency,
        "walks": [
            {
                "name": walk.name,
                "seconds": walk.seconds,
                "repeats": walk.repeats,
                "checksum": walk.checksum,
            }
            for walk in walks
        ],
        "landing_api_breaks": list(LANDING_API_BREAKS),
    }


def main(argv: list[str] | None = None) -> int:
    """Build a graph, construct the CSR sidecar, and print the spike table."""
    parser = argparse.ArgumentParser(description=_DESCRIPTION)
    parser.add_argument(
        "--preset",
        choices=("taco", "if-heavy", "lic-dsf", "synthetic"),
        default="taco",
        help="Workbook/graph source (default: taco_patterns.xlsx)",
    )
    parser.add_argument("--workbook", type=Path, default=None)
    parser.add_argument("--targets", nargs="+", default=None)
    parser.add_argument("--load-values", action="store_true")
    parser.add_argument("--no-provenance", action="store_true")
    parser.add_argument("--use-cached-dynamic-refs", action="store_true")
    parser.add_argument("--synthetic-nodes", type=int, default=2000)
    parser.add_argument("--synthetic-degree", type=int, default=6)
    parser.add_argument(
        "--eval-order",
        action="store_true",
        help="Time DependencyGraph.evaluation_order (rebuilds dict adjacency)",
    )
    parser.add_argument(
        "--eval-order-max-nodes",
        type=int,
        default=5000,
        help="Auto-run evaluation_order when node count is at most this (0 disables)",
    )
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    source, graph = _build_graph(args)
    tracemalloc.start()
    csr = from_graph(graph)
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    include_eval_order = args.eval_order or (
        args.eval_order_max_nodes > 0 and len(graph) <= args.eval_order_max_nodes
    )
    walks = time_walks(graph, csr, include_eval_order=include_eval_order)
    if args.json:
        print(
            json.dumps(
                report_to_dict(source=source, graph=graph, csr=csr, build_peak=peak, walks=walks),
                indent=2,
            )
        )
    else:
        print(
            render_report(
                source=source,
                graph=graph,
                csr=csr,
                build_peak=peak,
                walks=walks,
                include_eval_order=include_eval_order,
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
