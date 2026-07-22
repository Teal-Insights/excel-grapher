#!/usr/bin/env python3
"""Codegen the LIC-DSF template graph into a local package under ``local-testing/``.

Loads the cached dependency graph for ``lic-dsf-template-2025-08-12.xlsm``
(same cache as ``examples/lic_dsf/extract_graph_cached.py``), exports Chart Data
targets via ``CodeGenerator.generate_modules``, and writes a package under
``local-testing/``:

- default: ``lic_dsf_2025_08_12/``
- ``--formula-groups``: ``lic_dsf_2025_08_12_formula_groups/``

Each package contains ``__init__.py``, ``api.py``, ``data.py``, ``runtime.py``,
and ``internals.py``.

Run from the repo root::

    uv run python local-testing/codegen_lic_dsf.py
    uv run python local-testing/codegen_lic_dsf.py --formula-groups
    uv run python local-testing/codegen_lic_dsf.py --formula-groups --hash-group-helpers
    uv run python local-testing/codegen_lic_dsf.py --no-cache

Requires the workbook under ``examples/lic_dsf/data/`` and (unless ``--no-cache``)
the graph cache under ``examples/lic_dsf/.cache/``.
"""

from __future__ import annotations

import argparse
import gzip
import json
import sys
import time
from pathlib import Path
from typing import cast

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from excel_grapher import CodeGenerator, create_dependency_graph, get_calc_settings
from excel_grapher.grapher.cache import (
    CacheValidationPolicy,
    GraphCacheMeta,
    build_graph_cache_meta,
    try_load_graph_cache,
)
from excel_grapher.grapher.formula_groups import coalesce_formula_groups
from excel_grapher.grapher.graph import DependencyGraph
from tests.integration.evaluator.utils.lic_dsf_chart_targets import (
    WORKBOOK_PATH,
    collect_chart_data_cell_keys,
)

_LIC_DSF_DIR = REPO_ROOT / "examples" / "lic_dsf"
DEFAULT_CACHE = _LIC_DSF_DIR / ".cache" / "lic-dsf-template-2025-08-12-dependency-graph.json.gz"
PACKAGE_NAME = "lic_dsf_2025_08_12"
PACKAGE_NAME_FORMULA_GROUPS = "lic_dsf_2025_08_12_formula_groups"
GRAPH_MAX_DEPTH = 50
GRAPH_LOAD_VALUES = True
GRAPH_USE_CACHED_DYNAMIC_REFS = True


def _default_out_dir(*, formula_groups: bool) -> Path:
    name = PACKAGE_NAME_FORMULA_GROUPS if formula_groups else PACKAGE_NAME
    return Path(__file__).resolve().parent / name


def _package_name(*, formula_groups: bool) -> str:
    return PACKAGE_NAME_FORMULA_GROUPS if formula_groups else PACKAGE_NAME


def _load_cached_graph(cache_path: Path, expected_meta: GraphCacheMeta) -> DependencyGraph | None:
    if not cache_path.is_file():
        return None
    # Prefer portable policy so this script works even if workbook mtime drifted.
    graph = try_load_graph_cache(
        cache_path,
        expected_meta=expected_meta,
        policy=CacheValidationPolicy.PORTABLE,
    )
    if graph is not None:
        return graph
    # Fall back to reading meta from the cache file itself (same format as regenerate_sample_viz).
    raw = cache_path.read_bytes()
    payload = (
        json.loads(gzip.decompress(raw).decode("utf-8"))
        if cache_path.suffix == ".gz" or cache_path.name.endswith(".json.gz")
        else json.loads(raw.decode("utf-8"))
    )
    meta = cast(GraphCacheMeta, payload["meta"])
    return try_load_graph_cache(
        cache_path,
        expected_meta=meta,
        policy=CacheValidationPolicy.PORTABLE,
    )


def _write_modules(module_dir: Path, files: dict[str, str]) -> None:
    module_dir.mkdir(parents=True, exist_ok=True)
    for filename, content in files.items():
        path = module_dir / filename
        path.write_text(content, encoding="utf-8")
        print(f"  wrote {path} ({len(content):,} bytes)")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=WORKBOOK_PATH,
        help="LIC-DSF workbook path",
    )
    parser.add_argument(
        "--cache",
        type=Path,
        default=DEFAULT_CACHE,
        help="Dependency-graph cache (.json / .json.gz)",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help=(
            "Directory for the generated package "
            f"(default: local-testing/{PACKAGE_NAME} or "
            f"local-testing/{PACKAGE_NAME_FORMULA_GROUPS} with --formula-groups)"
        ),
    )
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Force create_dependency_graph instead of loading the cache",
    )
    parser.add_argument(
        "--formula-groups",
        action="store_true",
        help=(
            "Run coalesce_formula_groups before codegen and write to "
            f"{PACKAGE_NAME_FORMULA_GROUPS}/"
        ),
    )
    parser.add_argument(
        "--hash-group-helpers",
        action="store_true",
        help=(
            "Emit compact hashed `_group_<sha1[:12]>` helper names "
            "(smaller internals.py for large multi-area groups)"
        ),
    )
    args = parser.parse_args()

    workbook = Path(args.workbook)
    if not workbook.is_file():
        raise SystemExit(f"Workbook not found: {workbook}")

    package_name = _package_name(formula_groups=args.formula_groups)
    out_dir = Path(args.out) if args.out is not None else _default_out_dir(
        formula_groups=args.formula_groups
    )

    targets = collect_chart_data_cell_keys()
    print(f"workbook: {workbook}")
    print(f"targets:  {len(targets)} Chart Data export cells")
    print(f"package:  {package_name}")
    print(f"out:      {out_dir}")

    expected_meta = build_graph_cache_meta(
        workbook,
        targets,
        extraction_params={
            "schema": 3,
            "max_depth": GRAPH_MAX_DEPTH,
            "load_values": GRAPH_LOAD_VALUES,
            "use_cached_dynamic_refs": GRAPH_USE_CACHED_DYNAMIC_REFS,
        },
    )

    graph: DependencyGraph | None = None
    if not args.no_cache:
        print(f"\nLoading graph cache: {args.cache}")
        t0 = time.perf_counter()
        graph = _load_cached_graph(args.cache, expected_meta)
        if graph is not None:
            print(f"  loaded in {time.perf_counter() - t0:.1f}s  ({len(graph)} nodes)")
        else:
            print("  cache miss / unreadable; will rebuild")

    if graph is None:
        print("\nBuilding dependency graph (this can take several minutes)...")
        t0 = time.perf_counter()
        graph = create_dependency_graph(
            workbook,
            targets,
            load_values=GRAPH_LOAD_VALUES,
            max_depth=GRAPH_MAX_DEPTH,
            use_cached_dynamic_refs=GRAPH_USE_CACHED_DYNAMIC_REFS,
        )
        print(f"  built in {time.perf_counter() - t0:.1f}s  ({len(graph)} nodes)")

    if args.formula_groups:
        print("\ncoalesce_formula_groups...")
        t0 = time.perf_counter()
        # Clone so we do not mutate a shared cache-backed object unexpectedly.
        working = graph._copy_for_projection()
        report = coalesce_formula_groups(working)
        print(
            f"  done in {time.perf_counter() - t0:.1f}s  "
            f"created={len(report.created_groups)}  "
            f"skipped={len(report.skipped_families)}  "
            f"nodes={len(working)}"
        )
        graph = working

    settings = get_calc_settings(workbook)
    print(
        f"\ncalc settings: iterate={settings.iterate_enabled} "
        f"count={settings.iterate_count} delta={settings.iterate_delta}"
    )

    print("\nCodeGenerator.generate_modules...")
    t0 = time.perf_counter()
    with CodeGenerator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
        hash_group_helper_names=args.hash_group_helpers,
    ) as gen:
        modules = gen.generate_modules(targets)
    print(f"  generated in {time.perf_counter() - t0:.1f}s  ({len(modules)} files)")

    print(f"\nWriting package to {out_dir}")
    _write_modules(out_dir, modules)
    print(f"\nDone. Import with:\n  from {package_name} import compute_all, make_context")


if __name__ == "__main__":
    main()
