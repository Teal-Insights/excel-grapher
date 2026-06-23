#!/usr/bin/env python3
"""Run Sprint 0 operator array benchmarks and optionally refresh baseline JSON."""

from __future__ import annotations

import argparse
from pathlib import Path

from tests.bench.operators_bench import (
    collect_baseline,
    write_baseline_document,
)

DEFAULT_BASELINE_PATH = (
    Path(__file__).resolve().parents[1]
    / "tests"
    / "fixtures"
    / "operators_baseline"
    / "baseline.json"
)


def _print_results(results: list) -> None:
    print(f"{'workload':<32} {'cells':>8} {'elapsed_s':>12} {'cells_per_sec':>14}")
    print("-" * 70)
    for result in results:
        print(
            f"{result.name:<32} {result.cell_count:>8} "
            f"{result.elapsed_sec:>12.6f} {result.cells_per_sec:>14.2f}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--write-baseline",
        action="store_true",
        help="Write results to tests/fixtures/operators_baseline/baseline.json",
    )
    parser.add_argument(
        "--baseline-path",
        type=Path,
        default=DEFAULT_BASELINE_PATH,
        help="Baseline JSON path (used with --write-baseline)",
    )
    parser.add_argument("--warmup", type=int, default=2, help="Warmup rounds per workload")
    parser.add_argument("--rounds", type=int, default=5, help="Timed rounds per workload")
    args = parser.parse_args()

    results = collect_baseline(warmup_rounds=args.warmup, bench_rounds=args.rounds)
    _print_results(results)

    if args.write_baseline:
        write_baseline_document(args.baseline_path, results)
        print(f"\nWrote baseline to {args.baseline_path}")


if __name__ == "__main__":
    main()
