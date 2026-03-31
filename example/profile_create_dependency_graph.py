#!/usr/bin/env python3
"""
Profile :func:`excel_grapher.create_dependency_graph` with the same workbook, targets,
and dynamic-ref settings as ``map_lic_dsf_indicators.py`` (and ``compare_lic_graph_compression.py``).

Run from the repository root::

    uv run python example/profile_create_dependency_graph.py \\
        --cprofile-out /tmp/lic_graph.prof --cprofile-print 40

Smaller / faster iteration (shallow closure, few targets)::

    uv run python example/profile_create_dependency_graph.py --max-targets 30 --max-depth 12 \\
        --cprofile-out /tmp/lic_graph.prof --cprofile-print 30

Inspect the binary profile interactively::

    uv run python -m pstats /tmp/lic_graph.prof
    # then: sort cumulative
    #        stats 50
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

import example.extract_graph_uncached as lic  # noqa: E402
from excel_grapher import DynamicRefConfig, create_dependency_graph  # noqa: E402
from excel_grapher.grapher.graph import DependencyGraph  # noqa: E402


def _collect_targets() -> list[str]:
    out: list[str] = []
    for entry in lic.EXPORT_RANGES:
        sheet_name, range_a1 = lic.parse_range_spec(entry["range_spec"])
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


def _build_kwargs(workbook: Path, *, max_depth: int):
    dynamic_refs: DynamicRefConfig | None = None
    if not lic.USE_CACHED_DYNAMIC_REFS:
        dynamic_refs = DynamicRefConfig.from_constraints_and_workbook(
            lic.LicDsfConstraints,
            workbook,
        )
    return dict(
        load_values=False,
        max_depth=max_depth,
        dynamic_refs=dynamic_refs,
        use_cached_dynamic_refs=lic.USE_CACHED_DYNAMIC_REFS,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
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
        help="Use only the first N target cells (smaller closure for faster profiling).",
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=50,
        metavar="D",
        help="BFS depth for create_dependency_graph (smaller = faster, incomplete graph).",
    )
    parser.add_argument(
        "--no-provenance",
        action="store_true",
        help="Build with capture_dependency_provenance=False (often much faster; different graph attrs).",
    )
    parser.add_argument(
        "--cprofile-out",
        type=Path,
        default=None,
        metavar="FILE",
        help="Write cProfile stats to this file (binary).",
    )
    parser.add_argument(
        "--cprofile-print",
        type=int,
        default=0,
        metavar="N",
        help="Print top N lines by cumulative time to stderr after the run.",
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

    kwargs = _build_kwargs(wp, max_depth=args.max_depth)
    capture_prov = not args.no_provenance

    print("Workbook:", wp)
    print("Target cells:", len(targets))
    print("USE_CACHED_DYNAMIC_REFS:", lic.USE_CACHED_DYNAMIC_REFS)
    print("max_depth:", args.max_depth)
    print("capture_dependency_provenance:", capture_prov)
    print()

    prof: cProfile.Profile | None = None
    if args.cprofile_out is not None:
        prof = cProfile.Profile()
        prof.enable()

    t0 = time.perf_counter()
    graph = create_dependency_graph(
        wp,
        targets,
        capture_dependency_provenance=capture_prov,
        **kwargs,
    )
    elapsed = time.perf_counter() - t0

    if prof is not None:
        prof.disable()
        args.cprofile_out.parent.mkdir(parents=True, exist_ok=True)
        prof.dump_stats(str(args.cprofile_out))
        print(f"cProfile stats written to {args.cprofile_out.resolve()}", file=sys.stderr)
        if args.cprofile_print > 0:
            pstats.Stats(prof).sort_stats(pstats.SortKey.CUMULATIVE).print_stats(
                args.cprofile_print
            )

    n, e = _graph_sizes(graph)
    print(f"Nodes: {n}")
    print(f"Edges: {e}")
    print(f"Elapsed: {_fmt_s(elapsed)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
