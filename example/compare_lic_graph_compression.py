#!/usr/bin/env python3
"""
Compare dependency graph size before and after identity-transit compression for the
LIC-DSF example (same targets and dynamic-ref settings as map_lic_dsf_indicators.py).

Run from the repository root:

    uv run python example/compare_lic_graph_compression.py

Fast profiling loop (subset of targets, compress phase only, print top callees)::

    uv run python example/compare_lic_graph_compression.py --max-targets 20 \\
        --cprofile-out /tmp/lic_compress.prof --profile-stage compress --cprofile-print 40

One graph build (default): provenance is required for compression; we report node and
edge counts before ``compress_identity_transits()`` and after.

Optional ``--dual-build`` runs a second full build without provenance first, so you can
confirm the uncompressed graph matches the pre-compression graph (and pay roughly twice
the build cost).

Each step prints its own elapsed time plus a timing summary at the end.

Optional ``--workbook``, ``--max-targets``, and ``--cprofile-out`` / ``--cprofile-print`` /
``--profile-stage`` support dev workflows (smaller runs, profiling).
"""

from __future__ import annotations

import argparse
import cProfile
import pstats
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]

if str(REPO_ROOT / "example") not in sys.path:
    sys.path.insert(0, str(REPO_ROOT / "example"))

import map_lic_dsf_indicators as lic  # noqa: E402

from excel_grapher import DynamicRefConfig, create_dependency_graph  # noqa: E402
from excel_grapher.grapher.graph import DependencyGraph  # noqa: E402


def _collect_targets() -> list[str]:
    out: list[str] = []
    for _label, spec in lic.CHART_DATA_RANGES:
        sheet_name, range_a1 = lic.parse_range_spec(spec)
        out.extend(lic.cells_in_range(sheet_name, range_a1))
    return out


def _graph_sizes(g: DependencyGraph) -> tuple[int, int]:
    nodes = len(g)
    edges = sum(len(g.dependencies(k)) for k in g)
    return nodes, edges


def _fmt_s(seconds: float) -> str:
    if seconds >= 120:
        m, s = divmod(seconds, 60)
        return f"{int(m)}m {s:.1f}s"
    return f"{seconds:.2f}s"


def _build_kwargs(workbook: Path):
    dynamic_refs: DynamicRefConfig | None = None
    if not lic.USE_CACHED_DYNAMIC_REFS:
        dynamic_refs = DynamicRefConfig.from_constraints_and_workbook(
            lic.LicDsfConstraints,
            workbook,
        )
    return dict(
        load_values=False,
        max_depth=50,
        dynamic_refs=dynamic_refs,
        use_cached_dynamic_refs=lic.USE_CACHED_DYNAMIC_REFS,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--dual-build",
        action="store_true",
        help="Also build once without provenance (slow); node/edge counts should match "
        "the pre-compression graph.",
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=None,
        help="Override workbook path (default: map_lic_dsf_indicators.WORKBOOK_PATH).",
    )
    parser.add_argument(
        "--max-targets",
        type=int,
        default=None,
        metavar="N",
        help="Use only the first N target cells (faster dev loop for profiling).",
    )
    parser.add_argument(
        "--cprofile-out",
        type=Path,
        default=None,
        metavar="FILE",
        help="Write cProfile stats to this file (binary). Inspect with: "
        "uv run python -m pstats FILE",
    )
    parser.add_argument(
        "--cprofile-print",
        type=int,
        default=0,
        metavar="N",
        help="After profiling, print top N lines by cumulative time (stderr).",
    )
    parser.add_argument(
        "--profile-stage",
        choices=("all", "build", "compress"),
        default="all",
        help="Which phase to include under cProfile (default: all). Ignored without "
        "--cprofile-out.",
    )
    args = parser.parse_args()

    wp = (args.workbook if args.workbook is not None else REPO_ROOT / lic.WORKBOOK_PATH).resolve()
    if not wp.exists():
        print(f"Workbook not found: {wp}", file=sys.stderr)
        return 1

    targets = _collect_targets()
    if args.max_targets is not None:
        targets = targets[: max(0, args.max_targets)]
    if not targets:
        print("No target cells.", file=sys.stderr)
        return 1

    kwargs = _build_kwargs(wp)
    prof: cProfile.Profile | None = None
    if args.cprofile_out is not None:
        prof = cProfile.Profile()

    t_script0 = time.perf_counter()
    timings: list[tuple[str, float]] = []

    print("Workbook:", wp)
    print("Target cells:", len(targets))
    print("USE_CACHED_DYNAMIC_REFS:", lic.USE_CACHED_DYNAMIC_REFS)
    print()

    if args.dual_build:
        print("--- Baseline: no provenance (not compressible) ---")
        if prof is not None and args.profile_stage == "all":
            prof.enable()
        t0 = time.perf_counter()
        g0 = create_dependency_graph(wp, targets, capture_dependency_provenance=False, **kwargs)
        dt = time.perf_counter() - t0
        timings.append(("Baseline graph build (no provenance)", dt))
        n0, e0 = _graph_sizes(g0)
        print(f"  Nodes: {n0}")
        print(f"  Edges: {e0}")
        print(f"  Elapsed: {_fmt_s(dt)}")
        print()

    print("--- With provenance: before compress_identity_transits() ---")
    if prof is not None and args.profile_stage in ("all", "build"):
        prof.enable()
    t0 = time.perf_counter()
    g = create_dependency_graph(wp, targets, capture_dependency_provenance=True, **kwargs)
    dt_build = time.perf_counter() - t0
    timings.append(("Graph build (with provenance)", dt_build))
    n_before, e_before = _graph_sizes(g)
    print(f"  Nodes: {n_before}")
    print(f"  Edges: {e_before}")
    print(f"  Elapsed: {_fmt_s(dt_build)}")
    print()
    if prof is not None and args.profile_stage == "build":
        prof.disable()

    print("--- After compress_identity_transits() ---")
    # For "all", profiler stays on from the build section; only enable here for "compress" alone.
    if prof is not None and args.profile_stage == "compress":
        prof.enable()
    t0 = time.perf_counter()
    removed = g.compress_identity_transits()
    dt_comp = time.perf_counter() - t0
    timings.append(("compress_identity_transits()", dt_comp))
    n_after, e_after = _graph_sizes(g)
    print(f"  Nodes: {n_after}")
    print(f"  Edges: {e_after}")
    print(f"  Removed transit keys: {len(removed)}")
    print(f"  Elapsed: {_fmt_s(dt_comp)}")
    print()

    if n_before:
        pct = 100.0 * (1.0 - n_after / n_before)
        print(f"Node reduction: {pct:.4f}%  ({n_before} -> {n_after})")
    if e_before:
        pct_e = 100.0 * (1.0 - e_after / e_before)
        print(f"Edge reduction: {pct_e:.4f}%  ({e_before} -> {e_after})")

    t_total = time.perf_counter() - t_script0
    print("--- Timing summary ---")
    for label, dt in timings:
        print(f"  {label}: {_fmt_s(dt)}")
    print(f"  Total (measured steps): {_fmt_s(sum(dt for _, dt in timings))}")
    print(f"  Wall time (script): {_fmt_s(t_total)}")
    print()

    if prof is not None:
        prof.disable()
        args.cprofile_out.parent.mkdir(parents=True, exist_ok=True)
        prof.dump_stats(str(args.cprofile_out))
        print(f"cProfile stats written to {args.cprofile_out.resolve()}", file=sys.stderr)
        if args.cprofile_print > 0:
            stats = pstats.Stats(prof).sort_stats(pstats.SortKey.CUMULATIVE)
            stats.print_stats(args.cprofile_print)

    if args.dual_build:
        print()
        if n0 != n_before:
            print(
                f"Note: baseline nodes ({n0}) != pre-compress nodes ({n_before}); "
                "investigate if unexpected.",
                file=sys.stderr,
            )
        elif e0 != e_before:
            print(
                f"Note: baseline edges ({e0}) != pre-compress edges ({e_before}); "
                "investigate if unexpected.",
                file=sys.stderr,
            )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
