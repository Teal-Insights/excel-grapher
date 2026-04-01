from __future__ import annotations

import pytest

# ---------------------------------------------------------------------------
# Import the module under test
# ---------------------------------------------------------------------------
from example.diagnose_constraint_candidates import (
    TargetEntry,
    TraceCollector,
    TraceEvent,
    _format_seconds,
    _print_trace_summary,
    _stable_unique,
    build_subsets,
    select_entries,
)
from excel_grapher import DynamicRefTraceEvent

# ---------------------------------------------------------------------------
# _stable_unique
# ---------------------------------------------------------------------------


def test_stable_unique_preserves_order_and_removes_duplicates() -> None:
    result = _stable_unique(["b", "a", "b", "c", "a"])
    assert result == ("b", "a", "c")


def test_stable_unique_empty() -> None:
    assert _stable_unique([]) == ()


def test_stable_unique_no_duplicates() -> None:
    assert _stable_unique(["x", "y", "z"]) == ("x", "y", "z")


# ---------------------------------------------------------------------------
# _format_seconds
# ---------------------------------------------------------------------------


def test_format_seconds_under_60() -> None:
    assert _format_seconds(9.5) == "9.50s"


def test_format_seconds_exactly_60() -> None:
    result = _format_seconds(60.0)
    assert "1m" in result
    assert "0.00s" in result


def test_format_seconds_over_60() -> None:
    result = _format_seconds(90.0)
    assert "1m" in result
    assert "30.00s" in result


# ---------------------------------------------------------------------------
# select_entries
# ---------------------------------------------------------------------------


def _make_entries() -> list[TargetEntry]:
    return [
        TargetEntry(
            label="Alpha signals", range_spec="Sheet1!A1:A2", targets=("Sheet1!A1", "Sheet1!A2")
        ),
        TargetEntry(
            label="Beta output",
            range_spec="Sheet1!B1:B3",
            targets=("Sheet1!B1", "Sheet1!B2", "Sheet1!B3"),
        ),
        TargetEntry(label="Gamma", range_spec="Sheet1!C1", targets=("Sheet1!C1",)),
    ]


def test_select_entries_no_filter_returns_all() -> None:
    entries = _make_entries()
    result = select_entries(entries)
    assert result == entries


def test_select_entries_label_filter_case_insensitive() -> None:
    entries = _make_entries()
    result = select_entries(entries, label_filters=["alpha"])
    assert len(result) == 1
    assert result[0].label == "Alpha signals"


def test_select_entries_label_filter_partial_match() -> None:
    entries = _make_entries()
    result = select_entries(entries, label_filters=["signals"])
    assert len(result) == 1
    assert result[0].label == "Alpha signals"


def test_select_entries_entry_index() -> None:
    entries = _make_entries()
    result = select_entries(entries, entry_indexes=[0, 2])
    assert [e.label for e in result] == ["Alpha signals", "Gamma"]


def test_select_entries_first_entries_limit() -> None:
    entries = _make_entries()
    result = select_entries(entries, first_entries=2)
    assert len(result) == 2


def test_select_entries_first_entries_with_filter() -> None:
    entries = _make_entries()
    result = select_entries(entries, label_filters=["a"], first_entries=1)
    assert len(result) == 1


def test_select_entries_first_entries_zero() -> None:
    entries = _make_entries()
    result = select_entries(entries, first_entries=0)
    assert result == []


# ---------------------------------------------------------------------------
# build_subsets
# ---------------------------------------------------------------------------


def test_build_subsets_per_entry_creates_one_per_entry() -> None:
    entries = _make_entries()
    subsets = build_subsets(entries, per_entry=True)
    assert len(subsets) == 3
    assert subsets[0].name == "Alpha signals"
    assert subsets[0].targets == ("Sheet1!A1", "Sheet1!A2")


def test_build_subsets_per_entry_first_targets() -> None:
    entries = _make_entries()
    subsets = build_subsets(entries, per_entry=True, first_targets=1)
    assert subsets[0].targets == ("Sheet1!A1",)
    assert subsets[1].targets == ("Sheet1!B1",)


def test_build_subsets_combined_deduplicates() -> None:
    entries = [
        TargetEntry(label="A", range_spec="S!A1", targets=("S!A1", "S!A2")),
        TargetEntry(label="B", range_spec="S!A1:A3", targets=("S!A1", "S!A3")),
    ]
    subsets = build_subsets(entries, per_entry=False)
    assert len(subsets) == 1
    assert subsets[0].name == "combined"
    assert subsets[0].targets == ("S!A1", "S!A2", "S!A3")


def test_build_subsets_combined_first_targets() -> None:
    entries = _make_entries()
    subsets = build_subsets(entries, per_entry=False, first_targets=2)
    assert len(subsets[0].targets) == 2


def test_build_subsets_empty_entries_returns_empty() -> None:
    subsets = build_subsets([], per_entry=True)
    assert subsets == []


def test_build_subsets_skips_empty_per_entry_subsets() -> None:
    entries = [
        TargetEntry(label="A", range_spec="S!A1", targets=("S!A1",)),
        TargetEntry(label="B", range_spec="S!B1", targets=()),
    ]
    subsets = build_subsets(entries, per_entry=True)
    assert len(subsets) == 1
    assert subsets[0].name == "A"


# ---------------------------------------------------------------------------
# TraceCollector
# ---------------------------------------------------------------------------


def test_trace_collector_record_creates_event() -> None:
    trace = TraceCollector()
    trace.record(kind="infer", caller="fn", elapsed_s=0.5, detail="x=1")
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "infer"
    assert ev.caller == "fn"
    assert ev.elapsed_s == 0.5
    assert ev.detail == "x=1"
    assert ev.branch_estimate is None
    assert ev.context == ""


def test_trace_collector_context_label_captured_in_record() -> None:
    trace = TraceCollector()
    with trace.context("outer"):
        trace.record(kind="k", caller="fn", elapsed_s=0.0, detail="")
    assert trace._events[0].context == "outer"


def test_trace_collector_nested_context_joins_with_arrow() -> None:
    trace = TraceCollector()
    with trace.context("A"), trace.context("B"):
        trace.record(kind="k", caller="fn", elapsed_s=0.0, detail="")
    assert trace._events[0].context == "A > B"


def test_trace_collector_context_pops_after_exit() -> None:
    trace = TraceCollector()
    with trace.context("temporary"):
        pass
    trace.record(kind="k", caller="fn", elapsed_s=0.0, detail="")
    assert trace._events[0].context == ""


def test_trace_collector_context_pops_on_exception() -> None:
    trace = TraceCollector()
    with pytest.raises(ValueError), trace.context("ephemeral"):
        raise ValueError("boom")
    trace.record(kind="k", caller="fn", elapsed_s=0.0, detail="")
    assert trace._events[0].context == ""


def test_trace_collector_snapshot_and_since() -> None:
    trace = TraceCollector()
    assert trace.snapshot() == 0
    trace.record(kind="k1", caller="fn", elapsed_s=0.0, detail="")
    idx = trace.snapshot()
    assert idx == 1
    trace.record(kind="k2", caller="fn", elapsed_s=0.0, detail="")
    trace.record(kind="k3", caller="fn", elapsed_s=0.0, detail="")
    events = trace.since(idx)
    assert len(events) == 2
    assert events[0].kind == "k2"
    assert events[1].kind == "k3"


# ---------------------------------------------------------------------------
# TraceCollector.on_library_event
# ---------------------------------------------------------------------------


def test_on_library_event_maps_kind_and_name() -> None:
    trace = TraceCollector()
    lib_event = DynamicRefTraceEvent(
        kind="infer",
        name="infer_dynamic_offset_targets",
        elapsed_s=0.5,
        detail={"targets": 3, "formula": "=OFFSET(A1,1,0)", "current_sheet": "Sheet1"},
    )
    trace.on_library_event(lib_event)
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "infer"
    assert ev.caller == "infer_dynamic_offset_targets"
    assert ev.elapsed_s == 0.5
    assert "targets=3" in ev.detail


def test_on_library_event_extracts_branch_estimate() -> None:
    trace = TraceCollector()
    lib_event = DynamicRefTraceEvent(
        kind="build-domains",
        name="_build_domains",
        elapsed_s=0.1,
        detail={"refs": 2, "branch_estimate": 42},
    )
    trace.on_library_event(lib_event)
    ev = trace._events[0]
    assert ev.branch_estimate == 42


def test_on_library_event_extracts_count_as_branch_estimate() -> None:
    trace = TraceCollector()
    lib_event = DynamicRefTraceEvent(
        kind="offset-scalar-wide",
        name="_infer_offset_scalar_domains",
        elapsed_s=0.0,
        detail={"expr": "B1", "count": 20},
    )
    trace.on_library_event(lib_event)
    ev = trace._events[0]
    assert ev.branch_estimate == 20


def test_on_library_event_no_branch_estimate_when_absent() -> None:
    trace = TraceCollector()
    lib_event = DynamicRefTraceEvent(
        kind="offset-scalar-fallback",
        name="_infer_offset_scalar_domains",
        elapsed_s=0.0,
        detail={"expr": "Z99"},
    )
    trace.on_library_event(lib_event)
    ev = trace._events[0]
    assert ev.branch_estimate is None


def test_on_library_event_respects_context_stack() -> None:
    trace = TraceCollector()
    lib_event = DynamicRefTraceEvent(kind="infer", name="fn", elapsed_s=0.0)
    with trace.context("subset:Test"):
        trace.on_library_event(lib_event)
    ev = trace._events[0]
    assert ev.context == "subset:Test"


# ---------------------------------------------------------------------------
# _print_trace_summary
# ---------------------------------------------------------------------------


def test_print_trace_summary_empty_events(capsys) -> None:
    _print_trace_summary([], top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "no fallback-related events" in out


def test_print_trace_summary_shows_kind_counts(capsys) -> None:
    events = [
        TraceEvent(kind="fallback-domains", caller="fn", context="", elapsed_s=0.1, detail="x"),
        TraceEvent(kind="fallback-domains", caller="fn", context="", elapsed_s=0.2, detail="y"),
        TraceEvent(kind="infer", caller="fn", context="", elapsed_s=0.3, detail="z"),
    ]
    _print_trace_summary(events, top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "fallback-domains: 2" in out
    assert "infer: 1" in out


def test_print_trace_summary_filters_low_branch_events(capsys) -> None:
    events = [
        TraceEvent(
            kind="k", caller="fn", context="", elapsed_s=0.0, detail="low_detail", branch_estimate=4
        ),
        TraceEvent(
            kind="k",
            caller="fn",
            context="",
            elapsed_s=0.0,
            detail="high_detail",
            branch_estimate=100,
        ),
    ]
    _print_trace_summary(events, top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "high_detail" in out
    assert "low_detail" not in out


def test_print_trace_summary_falls_back_to_all_when_none_pass_filter(capsys) -> None:
    events = [
        TraceEvent(
            kind="k",
            caller="fn",
            context="",
            elapsed_s=0.0,
            detail="small_detail",
            branch_estimate=2,
        ),
    ]
    _print_trace_summary(events, top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "small_detail" in out  # shown because fallback: nothing passed the filter


def test_print_trace_summary_none_branch_estimate_always_included(capsys) -> None:
    events = [
        TraceEvent(
            kind="k", caller="fn", context="", elapsed_s=0.0, detail="low_detail", branch_estimate=2
        ),
        TraceEvent(
            kind="err",
            caller="fn",
            context="",
            elapsed_s=0.0,
            detail="err_detail",
            branch_estimate=None,
        ),
    ]
    _print_trace_summary(events, top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "err_detail" in out  # None branch_estimate always passes the filter
    assert "low_detail" not in out


def test_print_trace_summary_respects_top_events_limit(capsys) -> None:
    events = [
        TraceEvent(
            kind="k", caller="fn", context="", elapsed_s=0.0, detail=f"ev{i}", branch_estimate=100
        )
        for i in range(10)
    ]
    _print_trace_summary(events, top_events=3, trace_min_branches=16)
    out = capsys.readouterr().out
    shown = sum(1 for i in range(10) if f"ev{i}" in out)
    assert shown == 3
