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

from excel_grapher.core.formula_ast import parse
from excel_grapher.core.formula_shape import (
    FormulaShapeSummary,
    fingerprint_formula_shape,
    summarize_formula_shapes,
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
        `(summary, normalized_formulas)` for cardinality + parse-warm timing.
    """
    from collections import Counter

    from excel_grapher.core.formula_ast import FormulaParseError
    from excel_grapher.core.formula_normalization import normalize_excel_formula

    normalized: list[str] = []
    shape_counter: Counter[str] = Counter()
    unparseable = 0
    for sheet_name, raw in sheet_formulas:
        try:
            nf = normalize_excel_formula(raw, sheet_name).strip()
        except Exception:
            unparseable += 1
            continue
        if not nf:
            continue
        try:
            shape = fingerprint_formula_shape(nf)
        except FormulaParseError:
            unparseable += 1
            continue
        normalized.append(nf)
        shape_counter[shape.shape_key] += 1

    summary = FormulaShapeSummary(
        formula_nodes=len(normalized),
        distinct_normalized_formulas=len(set(normalized)),
        distinct_shapes=len(shape_counter),
        unparseable=unparseable,
        shape_counts=tuple(shape_counter.most_common()),
    )
    return summary, normalized


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
    from excel_grapher.core.formula_ast import FormulaParseError

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


def render_report(
    summary: FormulaShapeSummary,
    *,
    parse_times: dict[str, float] | None = None,
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

    if args.json:
        payload: dict[str, object] = {
            "workbook": str(args.workbook),
            "targets": targets_label,
            "summary": summary.to_dict(),
        }
        if parse_times is not None:
            payload["parse_warm"] = parse_times
        print(json.dumps(payload, indent=2))
    else:
        print(f"workbook: {args.workbook}")
        print(f"targets:  {targets_label}")
        print(render_report(summary, parse_times=parse_times, top_n=args.top))
    return 0


if __name__ == "__main__":
    sys.exit(main())
