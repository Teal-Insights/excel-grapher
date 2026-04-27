#!/usr/bin/env python3
"""Rebuild the LIC-DSF sample web viz from the cached dependency graph (defaults everywhere)."""

from __future__ import annotations

import argparse
import pickle
import time
from pathlib import Path

_DIR = Path(__file__).resolve().parent
OUT = _DIR / "data" / "lic-dsf-template-sample-exported-viz.html"
PKL = _DIR / ".cache" / "lic-dsf-template-2025-08-12-dependency-graph.pkl"

def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("-o", "--output", type=Path, default=OUT)
    p.add_argument("-c", "--cache", type=Path, default=PKL)
    p.add_argument("pickle", nargs="?", type=Path, help="Override cache path (positional)")
    args = p.parse_args()
    pkl = args.pickle or args.cache

    if not pkl.is_file():
        raise SystemExit(f"Missing {pkl}")

    from excel_grapher.exporter import to_web_viz_payload
    from excel_grapher.grapher.export import to_networkx
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.lightweight_viz import write_web_viz_html

    t0 = time.perf_counter()
    with pkl.open("rb") as f:
        blob = pickle.load(f)
    if not (isinstance(blob, tuple) and len(blob) == 2) or not isinstance(
        (g := blob[1]), DependencyGraph
    ):
        raise SystemExit("expected (meta, DependencyGraph) pickle")
    t1 = time.perf_counter()

    # All defaults: stratified_multipartite layout, module overlay, etc.
    payload = to_web_viz_payload(to_networkx(g))
    t2 = time.perf_counter()

    write_web_viz_html(
        payload,
        args.output,
        title="LIC-DSF Template dependency graph",
        data_mode="inline",
    )
    print(
        f"pickle+load {t1 - t0:.2f}s  build {t2 - t1:.2f}s  "
        f"nodes={payload.core.stats.node_count}  -> {args.output.resolve()}",
        flush=True,
    )

if __name__ == "__main__":
    main()