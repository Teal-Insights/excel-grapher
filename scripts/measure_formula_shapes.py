#!/usr/bin/env python3
"""Measure distinct normalized formulas vs punched AST shapes (#517 validate).

Reports the go/no-go cardinality ratio from GitHub #517: if punched shapes are
not substantially fewer than distinct `normalized_formula` strings, shape
interning mostly rediscovers the existing string-keyed AST cache (#337).

Also times an optimistic parse-warm bound: parse every distinct formula string
versus parse one representative formula per shape.

Usage:
    uv run python scripts/measure_formula_shapes.py
    uv run python scripts/measure_formula_shapes.py --workbook book.xlsx --targets 'Sheet1!A1'
    uv run python scripts/measure_formula_shapes.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from collections.abc import Sequence
from pathlib import Path

from excel_grapher.core.formula_ast import FormulaParseError, parse
from excel_grapher.core.formula_normalization import normalize_excel_formula
from excel_grapher.core.formula_shape import (
    FormulaShapeSummary,
    fingerprint_formula_shape,
    summarize_formula_shapes,
    summarize_normalized_formulas,
)
from excel_grapher.grapher.graph import DependencyGraph

_DESCRIPTION = "Measure distinct formula strings vs punched AST shapes (#517)."

_REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = _REPO_ROOT / "examples" / "micro_workbooks" / "taco_patterns.xlsx"
DEFAULT_TARGETS = (
    "Patterns!D3:D7",
    "Patterns!F3:F7",
    "Patterns!H3:H7",
    "Patterns!K3:K7",
    "Patterns!P3:P7",
)


def _collect_normalized_formulas(graph: DependencyGraph) -> list[str]:
    formulas: list[str] = []
    for _, node in graph.formula_nodes():
        nf = node.normalized_formula
        if isinstance(nf, str) and nf.strip():
            formulas.append(nf.strip())
    return formulas


def _scan_workbook_formulas(workbook: Path) -> list[tuple[str, str]]:
    """Return `(sheet_name, raw_formula)` for every formula cell in `workbook`."""
    import fastpyxl

    pairs: list[tuple[str, str]] = []
    wb = fastpyxl.load_workbook(workbook, data_only=False, read_only=True)
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.startswith("="):
                        pairs.append((sheet_name, value))
    finally:
        wb.close()
    return pairs


def summarize_scanned_formula_shapes(
    sheet_formulas: Sequence[tuple[str, str]],
) -> tuple[FormulaShapeSummary, list[str]]:
    """Normalize+fingerprint workbook-scanned formulas.

    Used by `--scan-workbook` so large workbooks can be measured without a
    full dependency-graph extraction. Named ranges are not expanded (unlike
    graph extraction), so counts are an approximation.

    Returns:
        `(summary, parseable_formulas)` for cardinality + parse-warm timing.
    """
    normalized: list[str] = []
    for sheet_name, raw in sheet_formulas:
        nf = normalize_excel_formula(raw, sheet_name).strip()
        if nf:
            normalized.append(nf)
    return summarize_normalized_formulas(normalized)


def measure_parse_warm_times(
    formulas: Sequence[str],
    *,
    repeats: int = 5,
) -> dict[str, float]:
    """Time string-keyed vs one-representative-per-shape parse warm paths.

    Shape-keyed timing is an optimistic lower bound: it assumes shapes are
    already identified and only one formula per shape needs parsing. Discovery
    still requires touching every distinct string unless identity comes from
    elsewhere (e.g. R1C1 interning).
    """
    distinct = sorted(set(formulas))
    if not distinct:
        return {
            "distinct_formulas": 0.0,
            "distinct_shapes": 0.0,
            "string_keyed_parse_s": 0.0,
            "shape_keyed_parse_s": 0.0,
            "repeats": float(repeats),
        }

    # Discover shapes once (not included in either timed warm path).
    reps_by_shape: dict[str, str] = {}
    parseable: list[str] = []
    for formula in distinct:
        try:
            shape = fingerprint_formula_shape(formula)
        except FormulaParseError:
            continue
        parseable.append(formula)
        reps_by_shape.setdefault(shape.shape_key, formula)
    distinct = parseable
    representatives = list(reps_by_shape.values())

    def _time_parse(items: Sequence[str]) -> float:
        best = float("inf")
        for _ in range(repeats):
            t0 = time.perf_counter()
            for formula in items:
                parse(formula)
            best = min(best, time.perf_counter() - t0)
        return best

    return {
        "distinct_formulas": float(len(distinct)),
        "distinct_shapes": float(len(representatives)),
        "string_keyed_parse_s": _time_parse(distinct),
        "shape_keyed_parse_s": _time_parse(representatives),
        "repeats": float(repeats),
    }


def measure_eval_times(
    graph: DependencyGraph,
    *,
    repeats: int = 3,
) -> dict[str, float]:
    """Time value-cache re-eval with string ASTs vs compiled shapes.

    Both paths parse/compile once in the first evaluate(); reported times are
    the subsequent evaluate() after clearing the value cache only (issue #517
    "eval time after invalidation").
    """
    from excel_grapher import FormulaEvaluator
    from excel_grapher.grapher.formula_shapes import warm_formula_shapes

    targets = list(graph.formula_keys())
    if not targets:
        return {
            "formula_cells": 0.0,
            "string_keyed_eval_s": 0.0,
            "shape_keyed_eval_s": 0.0,
            "repeats": float(repeats),
        }

    original = graph.formula_shapes
    table = original if original is not None else warm_formula_shapes(graph)

    def _reeval_s(*, with_shapes: bool) -> float:
        graph.formula_shapes = table if with_shapes else None
        best = float("inf")
        for _ in range(repeats):
            with FormulaEvaluator(graph, auto_detect_changes=False) as ev:
                ev.evaluate(targets)
                ev._cache.clear()
                t0 = time.perf_counter()
                ev.evaluate(targets)
                best = min(best, time.perf_counter() - t0)
        return best

    try:
        string_s = _reeval_s(with_shapes=False)
        shape_s = _reeval_s(with_shapes=True)
    finally:
        graph.formula_shapes = original
    return {
        "formula_cells": float(len(targets)),
        "string_keyed_eval_s": string_s,
        "shape_keyed_eval_s": shape_s,
        "repeats": float(repeats),
    }


def measure_codegen_sizes(
    graph: DependencyGraph,
    targets: Sequence[str],
) -> dict[str, float]:
    """Compare emitted LOC / helper count with and without interned shapes."""
    from excel_grapher.exporter.codegen import CodeGenerator
    from excel_grapher.grapher.formula_shapes import warm_formula_shapes

    original = graph.formula_shapes
    table = original if original is not None else warm_formula_shapes(graph)

    def _stats(code: str) -> dict[str, float]:
        return {
            "loc": float(code.count("\n") + 1),
            "shape_helpers": float(code.count("def _shape_")),
            "cell_functions": float(code.count("def cell_")),
        }

    try:
        graph.formula_shapes = None
        plain = _stats(CodeGenerator(graph).generate(list(targets)))
        graph.formula_shapes = table
        shaped = _stats(CodeGenerator(graph).generate(list(targets)))
    finally:
        graph.formula_shapes = original
    return {
        "string_keyed_loc": plain["loc"],
        "shape_keyed_loc": shaped["loc"],
        "shape_helpers": shaped["shape_helpers"],
        "cell_functions": shaped["cell_functions"],
    }


def render_report(
    summary: FormulaShapeSummary,
    *,
    parse_times: dict[str, float] | None = None,
    eval_times: dict[str, float] | None = None,
    codegen: dict[str, float] | None = None,
    top_n: int = 15,
) -> str:
    """Format a human-readable shape cardinality report."""
    lines = [
        f"formula nodes:                 {summary.formula_nodes:,}",
        f"distinct normalized formulas:  {summary.distinct_normalized_formulas:,}",
        f"distinct shapes:               {summary.distinct_shapes:,}",
        f"unparseable:                   {summary.unparseable:,}",
        f"shapes / formula strings:      {summary.shapes_per_formula_string:.4f}",
        f"mean instances / shape:        {summary.mean_instances_per_shape:.2f}",
    ]
    if parse_times is not None:
        string_s = parse_times["string_keyed_parse_s"]
        shape_s = parse_times["shape_keyed_parse_s"]
        speedup = (string_s / shape_s) if shape_s > 0 else float("inf")
        lines += [
            "",
            "parse warm (best of "
            f"{int(parse_times['repeats'])} repeats; shape path = 1 parse/shape):",
            f"  string-keyed:  {string_s * 1000:.3f} ms "
            f"({int(parse_times['distinct_formulas'])} parses)",
            f"  shape-keyed:   {shape_s * 1000:.3f} ms "
            f"({int(parse_times['distinct_shapes'])} parses)",
            f"  speedup bound: {speedup:.2f}x",
        ]
    if eval_times is not None:
        string_s = eval_times["string_keyed_eval_s"]
        shape_s = eval_times["shape_keyed_eval_s"]
        speedup = (string_s / shape_s) if shape_s > 0 else float("inf")
        lines += [
            "",
            f"eval after value-cache invalidation (best of {int(eval_times['repeats'])} repeats):",
            f"  string-keyed:  {string_s * 1000:.3f} ms",
            f"  shape-keyed:   {shape_s * 1000:.3f} ms",
            f"  speedup:       {speedup:.2f}x",
        ]
    if codegen is not None:
        lines += [
            "",
            "codegen:",
            f"  string-keyed LOC: {int(codegen['string_keyed_loc']):,}",
            f"  shape-keyed LOC:  {int(codegen['shape_keyed_loc']):,}",
            f"  shape helpers:    {int(codegen['shape_helpers']):,}",
            f"  cell functions:   {int(codegen['cell_functions']):,}",
        ]
    if summary.shape_counts:
        lines += ["", f"top shapes (up to {top_n}):"]
        for shape_key, count in summary.shape_counts[:top_n]:
            lines.append(f"  {count:>6,}  {shape_key}")
    return "\n".join(lines)


def _build_graph(args: argparse.Namespace) -> DependencyGraph:
    from excel_grapher import create_dependency_graph

    return create_dependency_graph(
        args.workbook,
        args.targets,
        load_values=args.load_values,
        capture_dependency_provenance=False,
        max_depth=50 if args.max_depth is None else args.max_depth,
        use_cached_dynamic_refs=args.use_cached_dynamic_refs,
    )


def main(argv: list[str] | None = None) -> int:
    """Build a graph and print formula-shape cardinality metrics."""
    parser = argparse.ArgumentParser(description=_DESCRIPTION)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help=f"Workbook to build the graph from (default: {DEFAULT_WORKBOOK.name})",
    )
    parser.add_argument(
        "--targets",
        nargs="+",
        default=list(DEFAULT_TARGETS),
        help="Sheet-qualified target cells or ranges",
    )
    parser.add_argument(
        "--load-values",
        action="store_true",
        help="Load cached Excel values while building the graph",
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=None,
        help="Optional dependency walk depth limit",
    )
    parser.add_argument(
        "--use-cached-dynamic-refs",
        action="store_true",
        help="Resolve dynamic refs from cached workbook values",
    )
    parser.add_argument(
        "--parse-repeats",
        type=int,
        default=5,
        help="Best-of-N repeats for parse-warm timings (default: 5)",
    )
    parser.add_argument(
        "--no-parse-timing",
        action="store_true",
        help="Skip parse-warm timings",
    )
    parser.add_argument(
        "--eval-repeats",
        type=int,
        default=3,
        help="Best-of-N repeats for eval-after-invalidation timings (default: 3)",
    )
    parser.add_argument(
        "--no-eval-timing",
        action="store_true",
        help="Skip evaluator timings (graph mode only)",
    )
    parser.add_argument(
        "--no-codegen",
        action="store_true",
        help="Skip codegen LOC / helper counts (graph mode only)",
    )
    parser.add_argument(
        "--scan-workbook",
        action="store_true",
        help=(
            "Scan every formula cell in the workbook (no dependency graph). "
            "Named ranges are not expanded."
        ),
    )
    parser.add_argument("--top", type=int, default=15, help="How many top shapes to print")
    parser.add_argument("--json", action="store_true", help="Emit JSON instead of text")
    args = parser.parse_args(argv)

    if not args.workbook.is_file():
        parser.error(f"workbook not found: {args.workbook}")

    graph: DependencyGraph | None = None
    if args.scan_workbook:
        scanned = _scan_workbook_formulas(args.workbook)
        summary, formulas = summarize_scanned_formula_shapes(scanned)
        targets_label: list[str] | str = "(scan-workbook: all formula cells)"
    else:
        graph = _build_graph(args)
        summary = summarize_formula_shapes(graph)
        formulas = _collect_normalized_formulas(graph)
        targets_label = list(args.targets)

    parse_times = (
        None
        if args.no_parse_timing
        else measure_parse_warm_times(formulas, repeats=args.parse_repeats)
    )
    eval_times = (
        None
        if graph is None or args.no_eval_timing
        else measure_eval_times(graph, repeats=args.eval_repeats)
    )
    codegen = (
        None
        if graph is None or args.no_codegen
        else measure_codegen_sizes(graph, list(args.targets))
    )

    if args.json:
        payload: dict[str, object] = {
            "workbook": str(args.workbook),
            "targets": targets_label,
            "summary": summary.to_dict(),
        }
        if parse_times is not None:
            payload["parse_warm"] = parse_times
        if eval_times is not None:
            payload["eval"] = eval_times
        if codegen is not None:
            payload["codegen"] = codegen
        print(json.dumps(payload, indent=2))
    else:
        print(f"workbook: {args.workbook}")
        print(f"targets:  {targets_label}")
        print(
            render_report(
                summary,
                parse_times=parse_times,
                eval_times=eval_times,
                codegen=codegen,
                top_n=args.top,
            )
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
