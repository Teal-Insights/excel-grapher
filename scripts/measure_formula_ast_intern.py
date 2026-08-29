#!/usr/bin/env python3
"""Measure formula-AST intern cost during extraction and JSON cache encode (#550).

Reports peak heap during `create_dependency_graph` (`tracemalloc`), intern-path
operation counts, and JSON cache artifact size. Steady-state
`scripts/measure_graph_memory.py` is a different measurement: intern-key strings
are discarded when extraction finishes.

Usage:
    uv run python scripts/measure_formula_ast_intern.py
    uv run python scripts/measure_formula_ast_intern.py --json
"""

from __future__ import annotations

import argparse
import json
import tempfile
import tracemalloc
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_grapher.core.formula_ast import AstNode
from excel_grapher.grapher.cache import dependency_graph_to_json
from excel_grapher.grapher.graph import DependencyGraph

_DESCRIPTION = "Measure formula-AST intern cost during extraction and cache encode (#550)."

_REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_FFV2 = _REPO_ROOT / "examples" / "micro_workbooks" / "ffv2.xlsx"
DEFAULT_TACO = _REPO_ROOT / "examples" / "micro_workbooks" / "taco_patterns.xlsx"
DEFAULT_TACO_TARGETS = (
    "Patterns!D3:D7",
    "Patterns!F3:F7",
    "Patterns!H3:H7",
    "Patterns!K3:K7",
    "Patterns!P3:P7",
)


@dataclass(frozen=True, slots=True)
class InternPathCounts:
    """Operation counts on the extraction intern path and resulting graph."""

    ast_to_json_calls: int
    json_dumps_intern_key_calls: int
    intern_hits: int
    intern_misses: int
    formula_nodes: int
    identity_distinct_trees: int
    equality_distinct_trees: int


@dataclass(frozen=True, slots=True)
class CacheArtifactStats:
    """JSON graph-cache payload size and formula-AST pool shape."""

    payload_bytes: int
    pool_entries: int
    node_ast_refs: int
    unique_ast_refs: int
    uses_formula_ast_id: bool
    uses_formula_ast_key: bool


@dataclass(frozen=True, slots=True)
class InternMeasurement:
    """One workbook/target measurement for #550."""

    name: str
    workbook: str
    targets: tuple[str, ...]
    peak_bytes: int
    current_bytes: int
    intern: InternPathCounts
    cache: CacheArtifactStats

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable view of this measurement."""
        return {
            "name": self.name,
            "workbook": self.workbook,
            "targets": list(self.targets),
            "peak_bytes": self.peak_bytes,
            "current_bytes": self.current_bytes,
            "intern": {
                "ast_to_json_calls": self.intern.ast_to_json_calls,
                "json_dumps_intern_key_calls": self.intern.json_dumps_intern_key_calls,
                "intern_hits": self.intern.intern_hits,
                "intern_misses": self.intern.intern_misses,
                "formula_nodes": self.intern.formula_nodes,
                "identity_distinct_trees": self.intern.identity_distinct_trees,
                "equality_distinct_trees": self.intern.equality_distinct_trees,
            },
            "cache": {
                "payload_bytes": self.cache.payload_bytes,
                "pool_entries": self.cache.pool_entries,
                "node_ast_refs": self.cache.node_ast_refs,
                "unique_ast_refs": self.cache.unique_ast_refs,
                "uses_formula_ast_id": self.cache.uses_formula_ast_id,
                "uses_formula_ast_key": self.cache.uses_formula_ast_key,
            },
        }


def _equality_distinct_count(trees: Sequence[AstNode]) -> int:
    """Return how many trees are distinct by `==` (works if trees are unhashable)."""
    unique: list[AstNode] = []
    for tree in trees:
        if tree not in unique:
            unique.append(tree)
    return len(unique)


def _json_intern_key_counts() -> tuple[int, int]:
    """Return intern-key JSON encodings on the extraction path.

    Schema 6 interned with `builder.ast_to_json` + `json.dumps`. After #550 those
    bindings are gone, so both counts are zero. A residual binding is reported as
    a nonzero sentinel so callers fail closed instead of under-counting.
    """
    import excel_grapher.grapher.builder as builder

    if hasattr(builder, "ast_to_json") or getattr(builder, "json", None) is not None:
        return 1, 1
    return 0, 0


def _intern_counts(
    graph: DependencyGraph, ast_to_json_calls: int, dumps_calls: int
) -> InternPathCounts:
    trees = [node.formula_ast for _, node in graph.formula_nodes() if node.formula_ast is not None]
    identity_distinct = len({id(tree) for tree in trees})
    equality_distinct = _equality_distinct_count(trees)
    return InternPathCounts(
        ast_to_json_calls=ast_to_json_calls,
        json_dumps_intern_key_calls=dumps_calls,
        intern_hits=len(trees) - identity_distinct,
        intern_misses=identity_distinct,
        formula_nodes=len(trees),
        identity_distinct_trees=identity_distinct,
        equality_distinct_trees=equality_distinct,
    )


def _cache_stats(graph: DependencyGraph) -> CacheArtifactStats:
    payload = dependency_graph_to_json(graph)
    encoded = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    pool = payload.get("formula_asts", [])
    pool_entries = len(pool) if isinstance(pool, (dict, list)) else 0
    refs: list[object] = []
    uses_id = False
    uses_key = False
    for node_payload in payload["nodes"]:
        if "formula_ast_id" in node_payload:
            uses_id = True
            refs.append(node_payload["formula_ast_id"])
        if "formula_ast_key" in node_payload:
            uses_key = True
            refs.append(node_payload["formula_ast_key"])
    unique_refs = len(set(refs))
    return CacheArtifactStats(
        payload_bytes=len(encoded),
        pool_entries=pool_entries,
        node_ast_refs=len(refs),
        unique_ast_refs=unique_refs,
        uses_formula_ast_id=uses_id,
        uses_formula_ast_key=uses_key,
    )


def measure_intern(
    workbook: Path,
    targets: Sequence[str],
    *,
    name: str,
    load_values: bool = False,
) -> InternMeasurement:
    """Build a graph for `workbook`/`targets` and return intern/cache measurements."""
    from excel_grapher import create_dependency_graph

    ast_calls, dumps_calls = _json_intern_key_counts()
    tracemalloc.start()
    try:
        graph = create_dependency_graph(workbook, list(targets), load_values=load_values)
        current, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()

    return InternMeasurement(
        name=name,
        workbook=str(workbook),
        targets=tuple(targets),
        peak_bytes=peak,
        current_bytes=current,
        intern=_intern_counts(graph, ast_calls, dumps_calls),
        cache=_cache_stats(graph),
    )


def _write_offset_workbook(path: Path) -> None:
    """Write `=A1*2` at B1 and B2 (same text, different relative offsets)."""
    import fastpyxl

    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["A2"].value = 20
    ws["B1"].value = "=A1*2"
    ws["B2"].value = "=A1*2"
    wb.save(path)
    wb.close()


def _write_large_autofill_workbook(path: Path, rows: int) -> None:
    """Write an autofill column of `=A{n}*2` so intern hits dominate."""
    import fastpyxl

    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in range(1, rows + 1):
        ws.cell(row, 1).value = float(row)
        ws.cell(row, 2).value = f"=A{row}*2"
    wb.save(path)
    wb.close()


def default_cases(tmp_dir: Path) -> list[tuple[str, Path, tuple[str, ...]]]:
    """Return named (label, workbook, targets) cases required by #550."""
    cases: list[tuple[str, Path, tuple[str, ...]]] = []
    if DEFAULT_FFV2.is_file():
        cases.append(("ffv2 B18:Q18", DEFAULT_FFV2, ("Sheet1!B18:Q18",)))
    if DEFAULT_TACO.is_file():
        cases.append(("taco_patterns", DEFAULT_TACO, DEFAULT_TACO_TARGETS))
    offset = tmp_dir / "same_text_different_offset.xlsx"
    _write_offset_workbook(offset)
    cases.append(("same-text different-offset =A1*2", offset, ("Sheet1!B1", "Sheet1!B2")))
    large = tmp_dir / "large_autofill.xlsx"
    _write_large_autofill_workbook(large, 400)
    cases.append(("large autofill 400x =An*2", large, ("Sheet1!B1:B400",)))
    return cases


def _kib(value: int) -> str:
    return f"{value / 1024:,.1f} KiB"


def render(measurements: Sequence[InternMeasurement]) -> str:
    """Return a human-readable table of intern measurements."""
    lines = [
        f"{'case':<36}{'peak':>12}{'cur':>12}{'to_json':>9}{'dumps':>8}"
        f"{'hits':>8}{'miss':>8}{'trees':>8}{'cache':>12}{'pool':>8}",
        "-" * 121,
    ]
    for item in measurements:
        lines.append(
            f"{item.name:<36}{item.peak_bytes:>12,}{item.current_bytes:>12,}"
            f"{item.intern.ast_to_json_calls:>9,}{item.intern.json_dumps_intern_key_calls:>8,}"
            f"{item.intern.intern_hits:>8,}{item.intern.intern_misses:>8,}"
            f"{item.intern.identity_distinct_trees:>8,}{item.cache.payload_bytes:>12,}"
            f"{item.cache.pool_entries:>8,}"
        )
    lines += [
        "",
        "peak/cur = tracemalloc around create_dependency_graph (bytes)",
        "to_json  = ast_to_json calls on the extraction intern path (must be 0 after #550)",
        "dumps    = json.dumps intern-key encodings on the extraction intern path (must be 0)",
        "hits     = formula nodes sharing an interned tree by identity",
        "miss     = identity-distinct interned trees (also intern_misses)",
        "cache    = json.dumps(dependency_graph_to_json(graph), compact) bytes",
        "pool     = formula_asts pool entries",
    ]
    for item in measurements:
        lines.append(
            f"  {item.name}: cache id={item.cache.uses_formula_ast_id} "
            f"key={item.cache.uses_formula_ast_key} "
            f"eq_distinct={item.intern.equality_distinct_trees} "
            f"peak={_kib(item.peak_bytes)}"
        )
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    """Run intern measurements and print a table or JSON."""
    parser = argparse.ArgumentParser(description=_DESCRIPTION)
    parser.add_argument("--json", action="store_true", help="Emit JSON instead of a table")
    parser.add_argument(
        "--workbook",
        type=Path,
        help="Measure a single workbook instead of the default #550 cases",
    )
    parser.add_argument("--targets", nargs="+", help="Targets for --workbook")
    args = parser.parse_args(argv)

    with tempfile.TemporaryDirectory(prefix="formula-ast-intern-") as raw_tmp:
        tmp_dir = Path(raw_tmp)
        if args.workbook is not None:
            if not args.workbook.is_file():
                parser.error(f"workbook not found: {args.workbook}")
            if not args.targets:
                parser.error("--targets is required with --workbook")
            cases = [("custom", args.workbook, tuple(args.targets))]
        else:
            cases = default_cases(tmp_dir)
        measurements = [
            measure_intern(workbook, targets, name=name) for name, workbook, targets in cases
        ]

    if args.json:
        print(json.dumps([item.to_dict() for item in measurements], indent=2))
    else:
        print(render(measurements))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
