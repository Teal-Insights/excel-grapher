#!/usr/bin/env python3
"""Rebuild the LIC-DSF sample web viz from the cached dependency graph (defaults everywhere)."""

from __future__ import annotations

import argparse
import gzip
import json
import time
from pathlib import Path
from typing import cast

_DIR = Path(__file__).resolve().parent
OUT = _DIR / "data" / "lic-dsf-template-sample-exported-viz.html"
DEFAULT_CACHE = _DIR / ".cache" / "lic-dsf-template-2025-08-12-dependency-graph.json.gz"


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("-o", "--output", type=Path, default=OUT)
    p.add_argument("-c", "--cache", type=Path, default=DEFAULT_CACHE)
    p.add_argument(
        "cache_path",
        nargs="?",
        type=Path,
        help="Override graph cache path (positional; .json or .json.gz; same format as excel_grapher.grapher.cache)",
    )
    args = p.parse_args()
    cache_path = args.cache_path or args.cache

    if not cache_path.is_file():
        raise SystemExit(f"Missing {cache_path}")

    from excel_grapher.exporter import to_web_viz_payload
    from excel_grapher.grapher.cache import (
        CacheValidationPolicy,
        GraphCacheMeta,
        try_load_graph_cache,
    )
    from excel_grapher.grapher.export import to_networkx
    from excel_grapher.grapher.lightweight_viz import write_web_viz_html

    t0 = time.perf_counter()
    raw = cache_path.read_bytes()
    if cache_path.suffix == ".gz":
        root = json.loads(gzip.decompress(raw).decode("utf-8"))
    else:
        root = json.loads(raw.decode("utf-8"))
    meta = cast(GraphCacheMeta, root["meta"])
    graph = try_load_graph_cache(
        cache_path,
        expected_meta=meta,
        policy=CacheValidationPolicy.PORTABLE,
    )
    if graph is None:
        raise SystemExit("invalid or unreadable graph cache JSON")
    t1 = time.perf_counter()

    # All defaults: stratified_multipartite layout, module overlay, etc.
    payload = to_web_viz_payload(to_networkx(graph))
    t2 = time.perf_counter()

    write_web_viz_html(
        payload,
        args.output,
        title="LIC-DSF Template dependency graph",
        data_mode="inline",
    )
    print(
        f"cache_json+load {t1 - t0:.2f}s  build {t2 - t1:.2f}s  "
        f"nodes={payload.core.stats.node_count}  -> {args.output.resolve()}",
        flush=True,
    )


if __name__ == "__main__":
    main()
