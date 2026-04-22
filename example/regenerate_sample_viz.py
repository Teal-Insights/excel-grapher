#!/usr/bin/env python3
"""
Refresh the LIC-DSF sample lightweight viz HTML.

Default (fast): re-embed the current package ``lightweight_viz_template.html`` into
``example/data/lic-dsf-template-sample-exported-viz.html`` while keeping the existing
inline ``window.__VIZ_DATA__`` payload (no graph rebuild). That embedded JSON may be an
older **flat** snapshot (legacy inline shape); the viewer still accepts it. To regenerate
current wire JSON from the cached graph, use ``--full``. Choose ``--mode core`` to
rebuild a core-only payload or ``--mode exporter`` to include module-inference overlays.
Use ``--layout force`` for export-time force-directed coordinates (slow on large graphs;
default for this script). Timings print for pickle load and viz payload build.

Use ``--full`` to rebuild from ``example/.cache/...-dependency-graph.pkl`` (slow, large
RAM; re-runs ``to_lightweight_viz`` and serializes ~tens of MB of JSON).

To profile where time goes (with timeout / SIGTERM-safe flush), use
``example/profile_lightweight_viz.py`` (see its docstring).
"""

from __future__ import annotations

import argparse
import pickle
import re
import sys
import time
from importlib import resources
from pathlib import Path
from typing import Literal

_EXAMPLE_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _EXAMPLE_DIR.parent
_DEFAULT_HTML = _EXAMPLE_DIR / "data" / "lic-dsf-template-sample-exported-viz.html"
_DEFAULT_CACHE = _EXAMPLE_DIR / ".cache" / "lic-dsf-template-2025-08-12-dependency-graph.pkl"


def _lic_dsf_export_targets() -> tuple[str, ...]:
    from example.extract_graph_cached import EXPORT_RANGES, cells_in_range, parse_range_spec

    targets: list[str] = []
    seen: set[str] = set()
    for entry in EXPORT_RANGES:
        sheet, a1 = parse_range_spec(entry["range_spec"])
        for key in cells_in_range(sheet, a1):
            if key in seen:
                continue
            seen.add(key)
            targets.append(key)
    return tuple(targets)


def _embedded_payload_version(sample_path: Path) -> int | None:
    """Best-effort parse of ``\"version\"`` from the ``__VIZ_DATA__`` bootstrap line."""
    with sample_path.open("r", encoding="utf-8") as f:
        for line in f:
            if "window.__VIZ_DATA__" in line and "window.__VIZ_DATA_URL__" not in line:
                m = re.search(r'"version"\s*:\s*(\d+)', line)
                return int(m.group(1)) if m else None
    return None


def _package_template() -> str:
    import excel_grapher.grapher.lightweight_viz as lv

    pkg = lv.__package__ or "excel_grapher.grapher"
    return (
        resources.files(pkg).joinpath("lightweight_viz_template.html").read_text(encoding="utf-8")
    )


def _extract_title(html_head: str) -> str:
    m = re.search(r"<title>([^<]*)</title>", html_head, re.I)
    return m.group(1).strip() if m else "Workbook dependency graph"


def _extract_bootstrap_lines(sample_path: Path) -> tuple[str, str]:
    """
    Return (bootstrap_line, sidecar_line) from the first inline script in the sample.
    """
    bootstrap_line = ""
    sidecar_line = ""
    in_script = False
    with sample_path.open("r", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if not in_script:
                if s == "<script>":
                    in_script = True
                continue
            if s == "</script>":
                break
            if "window.__VIZ_DATA__" in line:
                bootstrap_line = line.rstrip("\n")
            elif "window.__VIZ_DATA_URL__" in line:
                sidecar_line = line.rstrip("\n")
                break
    if not bootstrap_line or not sidecar_line:
        raise ValueError(
            f"Could not parse bootstrap from {sample_path}: "
            "expected window.__VIZ_DATA__ and window.__VIZ_DATA_URL__ in first <script> block."
        )
    return bootstrap_line, sidecar_line


def refresh_template_only(sample_html: Path) -> None:
    with sample_html.open("r", encoding="utf-8") as f:
        head = f.read(65536)
    title = _extract_title(head)

    boot, side = _extract_bootstrap_lines(sample_html)
    tpl = _package_template()
    out_html = (
        tpl.replace("__TITLE__", title)
        .replace("/*__BOOTSTRAP__*/", boot.strip())
        .replace("/*__SIDECAR__*/", side.strip())
    )
    sample_html.write_text(out_html, encoding="utf-8")


def full_rebuild(
    sample_html: Path,
    cache_pkl: Path,
    budget_mb: int,
    mode: Literal["exporter", "core"],
    *,
    exclude_guarded: bool = False,
    layout: Literal["bfs", "layered", "grid", "force"] = "force",
) -> None:
    sys.path.insert(0, str(_REPO_ROOT))
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.lightweight_viz import write_lightweight_viz_html

    t0 = time.perf_counter()
    with cache_pkl.open("rb") as f:
        blob = pickle.load(f)
    pickle_sec = time.perf_counter() - t0
    if not isinstance(blob, tuple) or len(blob) != 2:
        raise SystemExit("Pickle must be (meta, graph) tuple")
    _, graph = blob
    if not isinstance(graph, DependencyGraph):
        raise SystemExit("Pickle graph is not a DependencyGraph")

    t1 = time.perf_counter()
    if mode == "exporter":
        from excel_grapher.exporter.lightweight_viz import to_lightweight_viz

        payload = to_lightweight_viz(
            graph,
            layout_mode=None if layout == "bfs" else layout,
        )
    else:
        from excel_grapher.grapher.lightweight_viz import (
            VizLimits,
            assemble_lightweight_viz_payload,
            build_lightweight_viz_core,
        )

        core = build_lightweight_viz_core(
            graph,
            limits=VizLimits(),
            layout_mode=layout,
            include_guarded_edges=not exclude_guarded,
            bfs_seed_keys=_lic_dsf_export_targets(),
            exclude_unreachable_from_bfs=True,
        )
        payload = assemble_lightweight_viz_payload(core, [])
    build_sec = time.perf_counter() - t1

    n_nodes = payload.core.stats.node_count
    print(
        f"Timing: pickle_load={pickle_sec:.3f}s, viz_payload_build={build_sec:.3f}s "
        f"(nodes={n_nodes}, layout={layout!r})",
        flush=True,
    )

    write_lightweight_viz_html(
        payload,
        sample_html,
        title="LIC-DSF Template dependency graph",
        data_mode="inline",
        inline_size_budget_mb=budget_mb,
    )


def main() -> None:
    from excel_grapher.grapher.lightweight_viz import VIZ_PAYLOAD_VERSION

    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--full",
        action="store_true",
        help="Rebuild from graph pickle (slow); default is template-only refresh.",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=_DEFAULT_HTML,
        help=f"Output HTML path (default: {_DEFAULT_HTML})",
    )
    p.add_argument(
        "--cache",
        type=Path,
        default=_DEFAULT_CACHE,
        help=f"Graph pickle for --full (default: {_DEFAULT_CACHE})",
    )
    p.add_argument(
        "--inline-budget-mb",
        type=int,
        default=512,
        help="Inline JSON size budget for --full (default: 512)",
    )
    p.add_argument(
        "--mode",
        choices=("exporter", "core"),
        default="core",
        help=(
            "Payload mode for --full (default: core): 'core' builds a core-only payload; "
            "'exporter' includes module-inference overlays."
        ),
    )
    p.add_argument(
        "--layout",
        choices=("bfs", "layered", "grid", "force"),
        default="force",
        help=(
            "Core node placement for --full (default: force). "
            "'bfs' / 'layered' / 'grid' use rank-band layouts; 'force' runs export-time "
            "force-directed layout (can take minutes on large graphs)."
        ),
    )
    p.add_argument(
        "--exclude-guarded",
        action="store_true",
        help=(
            "Core mode only: exclude guarded edges from core payload; unreachable nodes from "
            "the BFS seed set are pruned automatically."
        ),
    )
    args = p.parse_args()
    out = args.output.resolve()
    if args.full:
        cache = args.cache.resolve()
        if not cache.is_file():
            raise SystemExit(f"Missing cache: {cache}")
        print(
            f"Full rebuild from pickle in {args.mode!r} mode, layout={args.layout!r} "
            "(this may take many minutes)...",
            flush=True,
        )
        full_rebuild(
            out,
            cache,
            args.inline_budget_mb,
            args.mode,
            exclude_guarded=args.exclude_guarded,
            layout=args.layout,
        )
        print(f"Wrote {out} ({out.stat().st_size // 1024 // 1024} MiB)", flush=True)
        return

    if not out.is_file():
        raise SystemExit(
            f"Missing {out}; run with --full once after extract_graph_cached.py, "
            "or create the sample HTML first."
        )
    ver = _embedded_payload_version(out)
    refresh_template_only(out)
    print(f"Refreshed package template into {out} (data payload unchanged).", flush=True)
    if ver is not None and ver != VIZ_PAYLOAD_VERSION:
        print(
            f"Note: embedded JSON wire version is still {ver} "
            f"(expected {VIZ_PAYLOAD_VERSION} for current wire output). "
            f"Template-only refresh does not rebuild data. "
            f"Run with --full to rewrite inline data from the pickle: "
            f"{args.cache}",
            flush=True,
        )


if __name__ == "__main__":
    main()
