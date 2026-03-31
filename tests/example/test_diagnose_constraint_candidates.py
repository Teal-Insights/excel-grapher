from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    EnumDomain,
    IntervalDomain,
    RealIntervalDomain,
)


# ---------------------------------------------------------------------------
# Import the module under test (must exist for GREEN phase)
# ---------------------------------------------------------------------------
from example.diagnose_constraint_candidates import (
    TargetEntry,
    TargetSubset,
    TraceCollector,
    TraceEvent,
    _cell_type_domain_summary,
    _domain_size_from_interval,
    _estimate_branch_product,
    _format_seconds,
    _print_trace_summary,
    _stable_unique,
    build_subsets,
    install_trace_hooks,
    select_entries,
)


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
# _domain_size_from_interval
# ---------------------------------------------------------------------------


def test_domain_size_known_interval() -> None:
    assert _domain_size_from_interval(IntervalDomain(min=1, max=5)) == 5


def test_domain_size_single_value() -> None:
    assert _domain_size_from_interval(IntervalDomain(min=3, max=3)) == 1


def test_domain_size_inverted_is_zero() -> None:
    assert _domain_size_from_interval(IntervalDomain(min=5, max=2)) == 0


def test_domain_size_none_interval() -> None:
    assert _domain_size_from_interval(None) is None


def test_domain_size_unbounded_min() -> None:
    assert _domain_size_from_interval(IntervalDomain(min=None, max=5)) is None


def test_domain_size_unbounded_max() -> None:
    assert _domain_size_from_interval(IntervalDomain(min=1, max=None)) is None


# ---------------------------------------------------------------------------
# _cell_type_domain_summary
# ---------------------------------------------------------------------------


def test_cell_type_domain_summary_none() -> None:
    size, desc = _cell_type_domain_summary(None)
    assert size is None
    assert desc == "missing"


def test_cell_type_domain_summary_enum() -> None:
    ct = CellType(kind=CellKind.STRING, enum=EnumDomain(values=frozenset(["a", "b", "c"])))
    size, desc = _cell_type_domain_summary(ct)
    assert size == 3
    assert "enum[3]" in desc


def test_cell_type_domain_summary_interval() -> None:
    ct = CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=4))
    size, desc = _cell_type_domain_summary(ct)
    assert size == 4
    assert "interval" in desc
    assert "1" in desc
    assert "4" in desc


def test_cell_type_domain_summary_real_interval() -> None:
    ct = CellType(kind=CellKind.NUMBER, real_interval=RealIntervalDomain(min=0.0, max=1.0))
    size, desc = _cell_type_domain_summary(ct)
    assert size is None
    assert "real" in desc


def test_cell_type_domain_summary_plain_kind() -> None:
    ct = CellType(kind=CellKind.ANY)
    size, desc = _cell_type_domain_summary(ct)
    assert size is None
    assert desc == CellKind.ANY.value


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
        TargetEntry(label="Alpha signals", range_spec="Sheet1!A1:A2", targets=("Sheet1!A1", "Sheet1!A2")),
        TargetEntry(label="Beta output", range_spec="Sheet1!B1:B3", targets=("Sheet1!B1", "Sheet1!B2", "Sheet1!B3")),
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
    with trace.context("A"):
        with trace.context("B"):
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
    with pytest.raises(ValueError):
        with trace.context("ephemeral"):
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
# _estimate_branch_product
# ---------------------------------------------------------------------------


def test_estimate_branch_product_empty_refs_returns_one() -> None:
    total, detail = _estimate_branch_product([], {}, lookup_cell_type=lambda env, addr: None)
    assert total == 1
    assert detail == []


def test_estimate_branch_product_all_enum_returns_product() -> None:
    ct_a = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset([1, 2, 3])))
    ct_b = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset([10, 20])))
    total, detail = _estimate_branch_product(
        ["A1", "B1"], {}, lookup_cell_type=lambda env, addr: ct_a if addr == "A1" else ct_b
    )
    assert total == 6  # 3 * 2
    assert len(detail) == 2


def test_estimate_branch_product_unknown_size_returns_none() -> None:
    ct_a = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset([1, 2])))
    # B1 lookup returns None (missing) -> known=False
    total, detail = _estimate_branch_product(
        ["A1", "B1"], {}, lookup_cell_type=lambda env, addr: ct_a if addr == "A1" else None
    )
    assert total is None
    assert len(detail) == 2


def test_estimate_branch_product_deduplicates_refs() -> None:
    ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset([1, 2])))
    total, detail = _estimate_branch_product(
        ["A1", "A1", "A1"], {}, lookup_cell_type=lambda env, addr: ct
    )
    assert total == 2  # A1 counted once
    assert len(detail) == 1


def test_estimate_branch_product_interval_known() -> None:
    ct = CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5))
    total, detail = _estimate_branch_product(
        ["X1"], {}, lookup_cell_type=lambda env, addr: ct
    )
    assert total == 5


# ---------------------------------------------------------------------------
# install_trace_hooks — restoration
# ---------------------------------------------------------------------------


def test_install_trace_hooks_restores_originals_after_exit() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    originals = {
        "infer_dynamic_offset_targets": dr.infer_dynamic_offset_targets,
        "expand_leaf_env_to_argument_env": dr.expand_leaf_env_to_argument_env,
        "_build_domains": dr._build_domains,
        "_build_value_domains": dr._build_value_domains,
        "_infer_offset_scalar_domains": dr._infer_offset_scalar_domains,
    }
    trace = TraceCollector()
    with install_trace_hooks(trace):
        assert dr.infer_dynamic_offset_targets is not originals["infer_dynamic_offset_targets"]
        assert dr._build_domains is not originals["_build_domains"]
    for name, original in originals.items():
        assert getattr(dr, name) is original


def test_install_trace_hooks_restores_on_body_exception() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    original_offset = dr.infer_dynamic_offset_targets
    trace = TraceCollector()
    with pytest.raises(RuntimeError, match="body error"):
        with install_trace_hooks(trace):
            raise RuntimeError("body error")
    assert dr.infer_dynamic_offset_targets is original_offset


# ---------------------------------------------------------------------------
# install_trace_hooks — _wrap_infer
# ---------------------------------------------------------------------------


def test_install_trace_hooks_wrap_infer_records_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value={"Sheet1!X1"})
    with patch.object(dr, "infer_dynamic_offset_targets", stub):
        with install_trace_hooks(trace):
            result = dr.infer_dynamic_offset_targets("=OFFSET(A1,1,0)", current_sheet="Sheet1")
    assert result == {"Sheet1!X1"}
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "infer"
    assert ev.caller == "infer_dynamic_offset_targets"
    assert ev.elapsed_s >= 0
    assert "targets=1" in ev.detail


def test_install_trace_hooks_wrap_infer_context_includes_formula_preview() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value=set())
    with patch.object(dr, "infer_dynamic_index_targets", stub):
        with install_trace_hooks(trace):
            dr.infer_dynamic_index_targets("=INDEX(A:A,1)", current_sheet="Data")
    # The infer event context should encode the function name, sheet, and formula preview
    ev = trace._events[0]
    assert "infer_dynamic_index_targets" in ev.context
    assert "Data" in ev.context


# ---------------------------------------------------------------------------
# install_trace_hooks — _wrap_expand
# ---------------------------------------------------------------------------


def test_install_trace_hooks_wrap_expand_records_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value={"Sheet1!A1": CellType(kind=CellKind.NUMBER)})
    with patch.object(dr, "expand_leaf_env_to_argument_env", stub):
        with install_trace_hooks(trace):
            result = dr.expand_leaf_env_to_argument_env({"Sheet1!A1", "Sheet1!B1"})
    assert result == {"Sheet1!A1": CellType(kind=CellKind.NUMBER)}
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "argument-env"
    assert ev.caller == "expand_leaf_env_to_argument_env"
    assert "argument_refs=2" in ev.detail
    assert "inferred_cells=1" in ev.detail


def test_install_trace_hooks_wrap_expand_records_error_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(side_effect=ValueError("bad env"))
    with patch.object(dr, "expand_leaf_env_to_argument_env", stub):
        with install_trace_hooks(trace):
            with pytest.raises(ValueError, match="bad env"):
                dr.expand_leaf_env_to_argument_env({"Sheet1!A1"})
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "argument-env-error"
    assert "ValueError" in ev.detail
    assert "bad env" in ev.detail


# ---------------------------------------------------------------------------
# install_trace_hooks — _wrap_domains
# ---------------------------------------------------------------------------


def test_install_trace_hooks_wrap_domains_records_fallback_domains_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value={})
    limits = MagicMock()
    with patch.object(dr, "_build_domains", stub):
        with install_trace_hooks(trace):
            result = dr._build_domains([], {}, limits)
    assert result == {}
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "fallback-domains"
    assert ev.branch_estimate == 1  # empty addrs -> product 1, all known
    assert "refs=0" in ev.detail


def test_install_trace_hooks_wrap_domains_records_error_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(side_effect=RuntimeError("too many branches"))
    limits = MagicMock()
    with patch.object(dr, "_build_domains", stub):
        with install_trace_hooks(trace):
            with pytest.raises(RuntimeError, match="too many branches"):
                dr._build_domains([], {}, limits)
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "fallback-domains-error"
    assert "RuntimeError" in ev.detail


def test_install_trace_hooks_wrap_value_domains_records_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value={})
    limits = MagicMock()
    with patch.object(dr, "_build_value_domains", stub):
        with install_trace_hooks(trace):
            dr._build_value_domains([], {}, limits)
    assert len(trace._events) == 1
    assert trace._events[0].kind == "fallback-value-domains"


# ---------------------------------------------------------------------------
# install_trace_hooks — _wrap_offset_scalar
# ---------------------------------------------------------------------------


def test_install_trace_hooks_wrap_offset_scalar_fallback_records_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value=None)  # None -> fallback
    fake_node = MagicMock()
    with patch.object(dr, "_infer_offset_scalar_domains", stub):
        with patch.object(dr, "_ast_to_expr_string", return_value="OFFSET(A1,1,0)"):
            with install_trace_hooks(trace):
                result = dr._infer_offset_scalar_domains(
                    fake_node, {}, MagicMock(), None, current_sheet="Sheet1"
                )
    assert result is None
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "offset-scalar-fallback"
    assert "OFFSET(A1,1,0)" in ev.detail


def test_install_trace_hooks_wrap_offset_scalar_wide_records_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value=list(range(20)))  # > 8 -> wide
    fake_node = MagicMock()
    with patch.object(dr, "_infer_offset_scalar_domains", stub):
        with patch.object(dr, "_ast_to_expr_string", return_value="IDX"):
            with install_trace_hooks(trace):
                result = dr._infer_offset_scalar_domains(
                    fake_node, {}, MagicMock(), None, current_sheet="Sheet1"
                )
    assert result == list(range(20))
    assert len(trace._events) == 1
    ev = trace._events[0]
    assert ev.kind == "offset-scalar-wide"
    assert ev.branch_estimate == 20


def test_install_trace_hooks_wrap_offset_scalar_small_result_records_no_event() -> None:
    import excel_grapher.grapher.dynamic_refs as dr

    trace = TraceCollector()
    stub = MagicMock(return_value=[1, 2, 3])  # <= 8 -> no event
    fake_node = MagicMock()
    with patch.object(dr, "_infer_offset_scalar_domains", stub):
        with patch.object(dr, "_ast_to_expr_string", return_value="IDX"):
            with install_trace_hooks(trace):
                result = dr._infer_offset_scalar_domains(
                    fake_node, {}, MagicMock(), None, current_sheet="Sheet1"
                )
    assert result == [1, 2, 3]
    assert len(trace._events) == 0


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
        TraceEvent(kind="k", caller="fn", context="", elapsed_s=0.0, detail="low_detail", branch_estimate=4),
        TraceEvent(kind="k", caller="fn", context="", elapsed_s=0.0, detail="high_detail", branch_estimate=100),
    ]
    _print_trace_summary(events, top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "high_detail" in out
    assert "low_detail" not in out


def test_print_trace_summary_falls_back_to_all_when_none_pass_filter(capsys) -> None:
    events = [
        TraceEvent(kind="k", caller="fn", context="", elapsed_s=0.0, detail="small_detail", branch_estimate=2),
    ]
    _print_trace_summary(events, top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "small_detail" in out  # shown because fallback: nothing passed the filter


def test_print_trace_summary_none_branch_estimate_always_included(capsys) -> None:
    events = [
        TraceEvent(kind="k", caller="fn", context="", elapsed_s=0.0, detail="low_detail", branch_estimate=2),
        TraceEvent(kind="err", caller="fn", context="", elapsed_s=0.0, detail="err_detail", branch_estimate=None),
    ]
    _print_trace_summary(events, top_events=5, trace_min_branches=16)
    out = capsys.readouterr().out
    assert "err_detail" in out  # None branch_estimate always passes the filter
    assert "low_detail" not in out


def test_print_trace_summary_respects_top_events_limit(capsys) -> None:
    events = [
        TraceEvent(kind="k", caller="fn", context="", elapsed_s=0.0, detail=f"ev{i}", branch_estimate=100)
        for i in range(10)
    ]
    _print_trace_summary(events, top_events=3, trace_min_branches=16)
    out = capsys.readouterr().out
    shown = sum(1 for i in range(10) if f"ev{i}" in out)
    assert shown == 3
