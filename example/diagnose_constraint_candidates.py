#!/usr/bin/env python3
"""Diagnose missing dynamic-ref cell type constraints on small LIC-DSF target subsets.

This script reuses the target list from ``example.extract_graph_cached`` but lets
you run ``list_dynamic_ref_constraint_candidates()`` on much smaller slices so
you can inspect where inference falls back to combinatorial enumeration.
"""

from __future__ import annotations

import argparse
import inspect
import sys
import time
from collections import Counter
from collections.abc import Iterable, Iterator, Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from excel_grapher import DynamicRefConfig, list_dynamic_ref_constraint_candidates  # noqa: E402
from excel_grapher.core.cell_types import CellType, IntervalDomain  # noqa: E402


@dataclass(frozen=True, slots=True)
class TargetEntry:
    label: str
    range_spec: str
    targets: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class TargetSubset:
    name: str
    targets: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class TraceEvent:
    kind: str
    caller: str
    context: str
    elapsed_s: float
    detail: str
    branch_estimate: int | None = None


class TraceCollector:
    """Collect lightweight trace events for dynamic-ref fallback analysis."""

    def __init__(self) -> None:
        self._events: list[TraceEvent] = []
        self._context_stack: list[str] = []

    @contextmanager
    def context(self, label: str) -> Iterator[None]:
        self._context_stack.append(label)
        try:
            yield
        finally:
            self._context_stack.pop()

    def record(
        self,
        *,
        kind: str,
        caller: str,
        elapsed_s: float,
        detail: str,
        branch_estimate: int | None = None,
    ) -> None:
        self._events.append(
            TraceEvent(
                kind=kind,
                caller=caller,
                context=" > ".join(self._context_stack),
                elapsed_s=elapsed_s,
                detail=detail,
                branch_estimate=branch_estimate,
            )
        )

    def snapshot(self) -> int:
        return len(self._events)

    def since(self, index: int) -> list[TraceEvent]:
        return self._events[index:]


def _stable_unique(values: Iterable[str]) -> tuple[str, ...]:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        out.append(value)
    return tuple(out)


def select_entries(
    entries: Sequence[TargetEntry],
    *,
    label_filters: Sequence[str] = (),
    entry_indexes: Sequence[int] = (),
    first_entries: int | None = None,
) -> list[TargetEntry]:
    if not label_filters and not entry_indexes:
        selected = list(entries)
    else:
        wanted_indexes = set(entry_indexes)
        lowered_filters = tuple(f.lower() for f in label_filters)
        selected = []
        for idx, entry in enumerate(entries):
            label_match = any(fragment in entry.label.lower() for fragment in lowered_filters)
            if idx in wanted_indexes or label_match:
                selected.append(entry)
    if first_entries is not None:
        selected = selected[: max(0, first_entries)]
    return selected


def build_subsets(
    entries: Sequence[TargetEntry], *, per_entry: bool, first_targets: int | None = None
) -> list[TargetSubset]:
    if per_entry:
        subsets = [
            TargetSubset(
                name=entry.label,
                targets=entry.targets[: max(0, first_targets)]
                if first_targets is not None
                else entry.targets,
            )
            for entry in entries
        ]
        return [subset for subset in subsets if subset.targets]

    combined = _stable_unique(target for entry in entries for target in entry.targets)
    if first_targets is not None:
        combined = combined[: max(0, first_targets)]
    return [TargetSubset(name="combined", targets=combined)] if combined else []


def load_target_entries() -> list[TargetEntry]:
    import example.extract_graph_cached as cached

    entries: list[TargetEntry] = []
    for entry in cached.EXPORT_RANGES:
        label = entry["label"]
        range_spec = entry["range_spec"]
        sheet_name, range_a1 = cached.parse_range_spec(range_spec)
        entries.append(
            TargetEntry(
                label=label,
                range_spec=range_spec,
                targets=tuple(cached.cells_in_range(sheet_name, range_a1)),
            )
        )
    return entries


def build_dynamic_ref_config(workbook: Path) -> DynamicRefConfig:
    import example.extract_graph_uncached as uncached

    return DynamicRefConfig.from_constraints_and_workbook(
        uncached.LicDsfConstraints,
        workbook,
    )


def _domain_size_from_interval(interval: IntervalDomain | None) -> int | None:
    if interval is None or interval.min is None or interval.max is None:
        return None
    lo = int(interval.min)
    hi = int(interval.max)
    if hi < lo:
        return 0
    return hi - lo + 1


def _cell_type_domain_summary(cell_type: CellType | None) -> tuple[int | None, str]:
    if cell_type is None:
        return None, "missing"
    if cell_type.enum is not None:
        return len(cell_type.enum.values), f"enum[{len(cell_type.enum.values)}]"
    if cell_type.interval is not None:
        size = _domain_size_from_interval(cell_type.interval)
        bounds = f"[{cell_type.interval.min},{cell_type.interval.max}]"
        return size, f"interval{bounds}"
    if cell_type.real_interval is not None:
        return None, f"real[{cell_type.real_interval.min},{cell_type.real_interval.max}]"
    return None, cell_type.kind.value


def _estimate_branch_product(
    addrs: Iterable[str],
    env: Any,
    *,
    lookup_cell_type,
) -> tuple[int | None, list[str]]:
    refs = sorted(set(addrs))
    detail: list[str] = []
    total = 1
    known = True
    for addr in refs:
        cell_type = lookup_cell_type(env, addr)
        size, summary = _cell_type_domain_summary(cell_type)
        detail.append(f"{addr}={summary}")
        if size is None:
            known = False
            continue
        total *= size
    return (total if known else None), detail


@contextmanager
def install_trace_hooks(trace: TraceCollector) -> Iterator[None]:
    import excel_grapher.grapher.dynamic_refs as dynamic_refs

    originals = {
        "infer_dynamic_offset_targets": dynamic_refs.infer_dynamic_offset_targets,
        "infer_dynamic_index_targets": dynamic_refs.infer_dynamic_index_targets,
        "infer_dynamic_indirect_targets": dynamic_refs.infer_dynamic_indirect_targets,
        "expand_leaf_env_to_argument_env": dynamic_refs.expand_leaf_env_to_argument_env,
        "_build_domains": dynamic_refs._build_domains,
        "_build_value_domains": dynamic_refs._build_value_domains,
        "_infer_offset_scalar_domains": dynamic_refs._infer_offset_scalar_domains,
    }

    def _caller_name() -> str:
        frame = inspect.currentframe()
        if frame is None or frame.f_back is None or frame.f_back.f_back is None:
            return "unknown"
        return frame.f_back.f_back.f_code.co_name

    def _wrap_infer(name: str, fn):
        def wrapped(formula: str, *args, current_sheet: str, **kwargs):
            preview = " ".join(formula.split())[:120]
            with trace.context(f"{name}:{current_sheet}:{preview}"):
                t0 = time.perf_counter()
                result = fn(formula, *args, current_sheet=current_sheet, **kwargs)
                trace.record(
                    kind="infer",
                    caller=name,
                    elapsed_s=time.perf_counter() - t0,
                    detail=f"targets={len(result)}",
                )
                return result

        return wrapped

    def _wrap_expand(fn):
        def wrapped(*args, **kwargs):
            argument_refs = args[0] if args else kwargs.get("argument_refs", set())
            t0 = time.perf_counter()
            try:
                result = fn(*args, **kwargs)
            except Exception as exc:
                trace.record(
                    kind="argument-env-error",
                    caller="expand_leaf_env_to_argument_env",
                    elapsed_s=time.perf_counter() - t0,
                    detail=f"argument_refs={len(argument_refs)} error={type(exc).__name__}: {exc}",
                )
                raise
            trace.record(
                kind="argument-env",
                caller="expand_leaf_env_to_argument_env",
                elapsed_s=time.perf_counter() - t0,
                detail=f"argument_refs={len(argument_refs)} inferred_cells={len(result)}",
            )
            return result

        return wrapped

    def _wrap_domains(kind: str, fn):
        def wrapped(addrs, env, limits):
            t0 = time.perf_counter()
            branch_estimate, detail = _estimate_branch_product(
                addrs,
                env,
                lookup_cell_type=dynamic_refs._lookup_cell_type,
            )
            detail_preview = "; ".join(detail[:8])
            if len(detail) > 8:
                detail_preview += f"; ... +{len(detail) - 8} more"
            try:
                result = fn(addrs, env, limits)
            except Exception as exc:
                trace.record(
                    kind=f"{kind}-error",
                    caller=_caller_name(),
                    elapsed_s=time.perf_counter() - t0,
                    branch_estimate=branch_estimate,
                    detail=f"refs={len(set(addrs))} {detail_preview} :: {type(exc).__name__}: {exc}",
                )
                raise
            trace.record(
                kind=kind,
                caller=_caller_name(),
                elapsed_s=time.perf_counter() - t0,
                branch_estimate=branch_estimate,
                detail=f"refs={len(set(addrs))} {detail_preview}",
            )
            return result

        return wrapped

    def _wrap_offset_scalar(fn):
        def wrapped(node, cell_type_env, limits, eval_context, *, current_sheet):
            expr = dynamic_refs._ast_to_expr_string(node)
            t0 = time.perf_counter()
            result = fn(
                node,
                cell_type_env,
                limits,
                eval_context,
                current_sheet=current_sheet,
            )
            if result is None:
                trace.record(
                    kind="offset-scalar-fallback",
                    caller=_caller_name(),
                    elapsed_s=time.perf_counter() - t0,
                    detail=f"{expr}",
                )
            elif len(result) > 8:
                trace.record(
                    kind="offset-scalar-wide",
                    caller=_caller_name(),
                    elapsed_s=time.perf_counter() - t0,
                    branch_estimate=len(result),
                    detail=f"{expr} -> {len(result)} values",
                )
            return result

        return wrapped

    dynamic_refs.infer_dynamic_offset_targets = _wrap_infer(
        "infer_dynamic_offset_targets", originals["infer_dynamic_offset_targets"]
    )
    dynamic_refs.infer_dynamic_index_targets = _wrap_infer(
        "infer_dynamic_index_targets", originals["infer_dynamic_index_targets"]
    )
    dynamic_refs.infer_dynamic_indirect_targets = _wrap_infer(
        "infer_dynamic_indirect_targets", originals["infer_dynamic_indirect_targets"]
    )
    dynamic_refs.expand_leaf_env_to_argument_env = _wrap_expand(
        originals["expand_leaf_env_to_argument_env"]
    )
    dynamic_refs._build_domains = _wrap_domains("fallback-domains", originals["_build_domains"])
    dynamic_refs._build_value_domains = _wrap_domains(
        "fallback-value-domains", originals["_build_value_domains"]
    )
    dynamic_refs._infer_offset_scalar_domains = _wrap_offset_scalar(
        originals["_infer_offset_scalar_domains"]
    )
    try:
        yield
    finally:
        dynamic_refs.infer_dynamic_offset_targets = originals["infer_dynamic_offset_targets"]
        dynamic_refs.infer_dynamic_index_targets = originals["infer_dynamic_index_targets"]
        dynamic_refs.infer_dynamic_indirect_targets = originals["infer_dynamic_indirect_targets"]
        dynamic_refs.expand_leaf_env_to_argument_env = originals["expand_leaf_env_to_argument_env"]
        dynamic_refs._build_domains = originals["_build_domains"]
        dynamic_refs._build_value_domains = originals["_build_value_domains"]
        dynamic_refs._infer_offset_scalar_domains = originals["_infer_offset_scalar_domains"]


def _format_seconds(seconds: float) -> str:
    if seconds >= 60:
        minutes, remainder = divmod(seconds, 60)
        return f"{int(minutes)}m {remainder:.2f}s"
    return f"{seconds:.2f}s"


def _print_trace_summary(
    events: Sequence[TraceEvent],
    *,
    top_events: int,
    trace_min_branches: int,
) -> None:
    if not events:
        print("  Trace: no fallback-related events captured.")
        return

    counts = Counter(event.kind for event in events)
    print("  Trace counts:")
    for kind, count in sorted(counts.items()):
        print(f"    {kind}: {count}")

    ranked = sorted(
        events,
        key=lambda event: (
            event.branch_estimate or 0,
            event.elapsed_s,
        ),
        reverse=True,
    )
    interesting = [
        event
        for event in ranked
        if event.branch_estimate is None or event.branch_estimate >= trace_min_branches
    ]
    if not interesting:
        interesting = ranked

    print(f"  Top {min(top_events, len(interesting))} trace events:")
    for event in interesting[:top_events]:
        branch_text = (
            f" branches~{event.branch_estimate}" if event.branch_estimate is not None else ""
        )
        print(
            "    "
            f"[{event.kind}] {event.caller}{branch_text} "
            f"elapsed={_format_seconds(event.elapsed_s)}"
        )
        if event.context:
            print(f"      context: {event.context}")
        print(f"      detail: {event.detail}")


def _run_subset_scan(
    workbook: Path,
    targets: Sequence[str],
    *,
    dynamic_refs: DynamicRefConfig,
    max_depth: int,
) -> tuple[list[str], float]:
    t0 = time.perf_counter()
    result = list_dynamic_ref_constraint_candidates(
        workbook,
        targets,
        dynamic_refs=dynamic_refs,
        max_depth=max_depth,
    )
    return result, time.perf_counter() - t0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path("example/data/lic-dsf-template-2025-08-12.xlsm"),
        help="Workbook to inspect.",
    )
    parser.add_argument(
        "--label-filter",
        action="append",
        default=[],
        help="Keep only target entries whose label contains this substring (repeatable).",
    )
    parser.add_argument(
        "--entry-index",
        action="append",
        default=[],
        type=int,
        help="Keep only these 0-based EXPORT_RANGES indexes (repeatable).",
    )
    parser.add_argument(
        "--first-entries",
        type=int,
        default=None,
        help="After filtering, keep only the first N entries.",
    )
    parser.add_argument(
        "--first-targets",
        type=int,
        default=None,
        help="Within each subset, keep only the first N targets.",
    )
    parser.add_argument(
        "--per-entry",
        action="store_true",
        help="Scan each selected export entry separately instead of as one combined subset.",
    )
    parser.add_argument(
        "--per-target-limit",
        type=int,
        default=0,
        help="Within each subset, also scan the first N targets individually for tighter localization.",
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=50,
        help="max_depth forwarded to list_dynamic_ref_constraint_candidates().",
    )
    parser.add_argument(
        "--top-events",
        type=int,
        default=8,
        help="Show at most N trace events per scan.",
    )
    parser.add_argument(
        "--trace-min-branches",
        type=int,
        default=16,
        help="Prefer trace events with at least this estimated branch count.",
    )
    parser.add_argument(
        "--max-candidates",
        type=int,
        default=20,
        help="Show at most N missing-constraint candidates per scan.",
    )
    args = parser.parse_args()

    workbook = (
        (REPO_ROOT / args.workbook).resolve() if not args.workbook.is_absolute() else args.workbook
    )
    if not workbook.exists():
        print(f"Workbook not found: {workbook}", file=sys.stderr)
        return 1

    entries = load_target_entries()
    selected_entries = select_entries(
        entries,
        label_filters=args.label_filter,
        entry_indexes=args.entry_index,
        first_entries=args.first_entries,
    )
    subsets = build_subsets(
        selected_entries,
        per_entry=args.per_entry,
        first_targets=args.first_targets,
    )
    if not subsets:
        print("No target subsets selected.", file=sys.stderr)
        return 1

    print(f"Workbook: {workbook}")
    print(f"Selected entries: {len(selected_entries)} / {len(entries)}")
    for entry in selected_entries[:10]:
        print(f"  - {entry.label}: {len(entry.targets)} targets")
    if len(selected_entries) > 10:
        print(f"  ... and {len(selected_entries) - 10} more")
    print()

    print("Building DynamicRefConfig from LIC-DSF constraints...")
    t_config = time.perf_counter()
    dynamic_refs = build_dynamic_ref_config(workbook)
    print(f"DynamicRefConfig ready in {_format_seconds(time.perf_counter() - t_config)}")
    print(f"Constraint env entries: {len(dynamic_refs.cell_type_env)}")
    print()

    trace = TraceCollector()
    with install_trace_hooks(trace):
        for subset in subsets:
            scan_start = trace.snapshot()
            with trace.context(f"subset:{subset.name}"):
                candidates, elapsed = _run_subset_scan(
                    workbook,
                    subset.targets,
                    dynamic_refs=dynamic_refs,
                    max_depth=args.max_depth,
                )
            events = trace.since(scan_start)
            print(f"Subset: {subset.name}")
            print(f"  Targets: {len(subset.targets)}")
            print(f"  Missing candidates: {len(candidates)}")
            if candidates:
                shown = candidates[: max(0, args.max_candidates)]
                print(f"  Candidate sample: {shown}")
                if len(candidates) > len(shown):
                    print(f"  ... and {len(candidates) - len(shown)} more")
            print(f"  Elapsed: {_format_seconds(elapsed)}")
            _print_trace_summary(
                events,
                top_events=args.top_events,
                trace_min_branches=args.trace_min_branches,
            )
            print()

            if args.per_target_limit > 0:
                for target in subset.targets[: args.per_target_limit]:
                    target_scan_start = trace.snapshot()
                    with trace.context(f"target:{target}"):
                        candidates, elapsed = _run_subset_scan(
                            workbook,
                            [target],
                            dynamic_refs=dynamic_refs,
                            max_depth=args.max_depth,
                        )
                    target_events = trace.since(target_scan_start)
                    print(f"Target: {target}")
                    print(f"  Missing candidates: {len(candidates)}")
                    if candidates:
                        shown = candidates[: max(0, args.max_candidates)]
                        print(f"  Candidate sample: {shown}")
                        if len(candidates) > len(shown):
                            print(f"  ... and {len(candidates) - len(shown)} more")
                    print(f"  Elapsed: {_format_seconds(elapsed)}")
                    _print_trace_summary(
                        target_events,
                        top_events=args.top_events,
                        trace_min_branches=args.trace_min_branches,
                    )
                    print()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
