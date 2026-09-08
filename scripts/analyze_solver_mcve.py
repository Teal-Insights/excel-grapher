"""Replay a may_cycle_solver_mcve fragment through cycle_report (#533)."""

from __future__ import annotations

import argparse
import time
from pathlib import Path

from excel_grapher.grapher.guard import evaluate_guard
from excel_grapher.grapher.solver_mcve import load_solver_mcve


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path, help="solver_manifest.json or .json.gz")
    args = parser.parse_args(argv)

    t0 = time.perf_counter()
    mcve = load_solver_mcve(args.manifest)
    print(f"loaded kind={mcve.kind} scc_index={mcve.scc_index} in {time.perf_counter() - t0:.2f}s")
    print(f"scc_members={len(mcve.scc_members)} intra_scc_edges={len(mcve.intra_scc_edges)}")
    print(f"guard_refs={len(mcve.guard_refs)} cone={len(mcve.guard_cone_outside_scc)}")
    print(f"leaf_constraints={len(mcve.leaf_constraints)}")

    values = {key: rec.value for key, rec in mcve.nodes.items()}
    dead = live = unknown = unguarded = 0
    for edge in mcve.intra_scc_edges:
        if edge.guard is None:
            unguarded += 1
            live += 1
            continue
        result = evaluate_guard(edge.guard, values)
        if result is False:
            dead += 1
        elif result is True:
            live += 1
        else:
            unknown += 1
            live += 1
    print(
        "cached-value guards: "
        f"true={live - unguarded} false={dead} unknown={unknown} unguarded={unguarded}"
    )

    t1 = time.perf_counter()
    graph = mcve.to_graph()
    print(f"graph nodes={len(graph)} in {time.perf_counter() - t1:.2f}s")

    t2 = time.perf_counter()
    report = graph.cycle_report()
    print(f"cycle_report in {time.perf_counter() - t2:.2f}s")
    print(f"must={report.has_must_cycles} may={report.has_may_cycles}")
    print(f"may_scc_count={len(report.may_cycles)}")
    if report.may_cycles:
        sizes = sorted((len(s) for s in report.may_cycles), reverse=True)
        print(f"may_scc_sizes={sizes[:8]}")
    if report.example_may_cycle_path:
        print(f"example_may_path_len={len(report.example_may_cycle_path)}")


if __name__ == "__main__":
    main()
