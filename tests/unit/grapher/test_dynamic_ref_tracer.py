"""Tests for the first-class dynamic-ref tracing infrastructure."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefTraceEvent,
    trace_dynamic_refs,
)


class TestDynamicRefTraceEvent:
    def test_fields(self) -> None:
        event = DynamicRefTraceEvent(
            kind="infer",
            name="infer_dynamic_offset_targets",
            elapsed_s=0.5,
            detail={"targets": 3},
        )
        assert event.kind == "infer"
        assert event.name == "infer_dynamic_offset_targets"
        assert event.elapsed_s == 0.5
        assert event.detail == {"targets": 3}

    def test_frozen(self) -> None:
        event = DynamicRefTraceEvent(kind="infer", name="test", elapsed_s=0.0, detail={})
        with pytest.raises(AttributeError):
            event.kind = "other"  # ty: ignore[invalid-assignment]

    def test_defaults(self) -> None:
        event = DynamicRefTraceEvent(kind="infer", name="test", elapsed_s=0.0)
        assert event.detail == {}


class TestTraceDynamicRefs:
    def test_context_manager_collects_events(self) -> None:
        """Emitting inside the context manager delivers events to the callback."""
        from excel_grapher.grapher.dynamic_refs import _emit_trace

        collected: list[DynamicRefTraceEvent] = []
        event = DynamicRefTraceEvent(kind="test", name="f", elapsed_s=0.0)

        with trace_dynamic_refs(collected.append):
            _emit_trace(event)

        assert collected == [event]

    def test_no_tracer_is_silent(self) -> None:
        """Emitting without an active tracer does not raise."""
        from excel_grapher.grapher.dynamic_refs import _emit_trace

        _emit_trace(DynamicRefTraceEvent(kind="test", name="f", elapsed_s=0.0))

    def test_nesting(self) -> None:
        """Inner context manager overrides; outer is restored after exit."""
        from excel_grapher.grapher.dynamic_refs import _emit_trace

        outer: list[DynamicRefTraceEvent] = []
        inner: list[DynamicRefTraceEvent] = []
        e1 = DynamicRefTraceEvent(kind="outer", name="f", elapsed_s=0.0)
        e2 = DynamicRefTraceEvent(kind="inner", name="f", elapsed_s=0.0)
        e3 = DynamicRefTraceEvent(kind="outer-again", name="f", elapsed_s=0.0)

        with trace_dynamic_refs(outer.append):
            _emit_trace(e1)
            with trace_dynamic_refs(inner.append):
                _emit_trace(e2)
            _emit_trace(e3)

        assert outer == [e1, e3]
        assert inner == [e2]

    def test_cleanup_on_exception(self) -> None:
        """Tracer is removed even when the body raises."""
        from excel_grapher.grapher.dynamic_refs import _emit_trace

        collected: list[DynamicRefTraceEvent] = []

        with pytest.raises(RuntimeError), trace_dynamic_refs(collected.append):
            _emit_trace(DynamicRefTraceEvent(kind="ok", name="f", elapsed_s=0.0))
            raise RuntimeError("boom")

        # After the context exits, no tracer should be active
        stray: list[DynamicRefTraceEvent] = []
        _emit_trace(DynamicRefTraceEvent(kind="stray", name="f", elapsed_s=0.0))
        assert stray == []
        assert len(collected) == 1


class TestTraceEmissions:
    """Verify that the 7 hook points emit trace events."""

    def _collect(self, fn, *args, **kwargs) -> list[DynamicRefTraceEvent]:
        import contextlib

        collected: list[DynamicRefTraceEvent] = []
        with trace_dynamic_refs(collected.append), contextlib.suppress(Exception):
            fn(*args, **kwargs)
        return collected

    def test_infer_offset_emits(self) -> None:
        """A formula with no OFFSET calls still emits an infer event."""
        from excel_grapher.grapher.dynamic_refs import infer_dynamic_offset_targets

        events = self._collect(
            infer_dynamic_offset_targets,
            "=1+2",
            current_sheet="Sheet1",
            cell_type_env={},
        )
        assert any(e.kind == "infer" and e.name == "infer_dynamic_offset_targets" for e in events)

    def test_infer_index_emits(self) -> None:
        from excel_grapher.grapher.dynamic_refs import infer_dynamic_index_targets

        events = self._collect(
            infer_dynamic_index_targets,
            "=1+2",
            current_sheet="Sheet1",
            cell_type_env={},
        )
        assert any(e.kind == "infer" and e.name == "infer_dynamic_index_targets" for e in events)

    def test_infer_indirect_emits(self) -> None:
        from excel_grapher.grapher.dynamic_refs import infer_dynamic_indirect_targets

        events = self._collect(
            infer_dynamic_indirect_targets,
            '=INDIRECT("Sheet1!A1")',
            current_sheet="Sheet1",
            cell_type_env={},
        )
        assert any(e.kind == "infer" and e.name == "infer_dynamic_indirect_targets" for e in events)

    def test_build_domains_emits(self) -> None:
        from excel_grapher.core.cell_types import CellKind, CellType, IntervalDomain
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            _build_domains,
        )

        env: dict[str, CellType] = {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=1, max=3),
            ),
        }
        events = self._collect(
            _build_domains,
            {"Sheet1!A1"},
            env,
            DynamicRefLimits(),
        )
        assert any(e.kind == "build-domains" for e in events)

    def test_build_domains_error_emits(self) -> None:
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            _build_domains,
        )

        events = self._collect(
            _build_domains,
            {"Sheet1!Z99"},
            {},
            DynamicRefLimits(),
        )
        assert any(e.kind == "build-domains-error" for e in events)

    def test_build_value_domains_emits(self) -> None:
        from excel_grapher.core.cell_types import CellKind, CellType, EnumDomain
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            _build_value_domains,
        )

        env: dict[str, CellType] = {
            "Sheet1!A1": CellType(
                kind=CellKind.STRING,
                enum=EnumDomain(values=frozenset({"hello", "world"})),
            ),
        }
        events = self._collect(
            _build_value_domains,
            {"Sheet1!A1"},
            env,
            DynamicRefLimits(),
        )
        assert any(e.kind == "build-value-domains" for e in events)

    def test_expand_env_emits(self) -> None:
        """expand_leaf_env_to_argument_env emits an expand-env event on success."""
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            expand_leaf_env_to_argument_env,
        )

        events = self._collect(
            expand_leaf_env_to_argument_env,
            set(),
            lambda addr: None,
            lambda f, s: set(),
            {},
            DynamicRefLimits(),
        )
        assert any(
            e.kind == "expand-env" and e.name == "expand_leaf_env_to_argument_env" for e in events
        )

    def test_expand_env_error_emits(self) -> None:
        """expand_leaf_env_to_argument_env emits an expand-env-error event on failure."""
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            expand_leaf_env_to_argument_env,
        )

        def bad_formula(addr: str) -> str:
            raise RuntimeError("boom")

        events = self._collect(
            expand_leaf_env_to_argument_env,
            {"Sheet1!A1"},
            bad_formula,
            lambda f, s: set(),
            {},
            DynamicRefLimits(),
        )
        assert any(
            e.kind == "expand-env-error" and e.name == "expand_leaf_env_to_argument_env"
            for e in events
        )

    def test_offset_scalar_fallback_emits(self) -> None:
        """When _infer_offset_scalar_domains returns None, it emits a fallback event."""
        from excel_grapher.core.formula_ast import CellRefNode
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            _infer_offset_scalar_domains,
        )

        # A cell ref with no domain in env -> fallback
        node = CellRefNode(address="Sheet1!Z99")
        events = self._collect(
            _infer_offset_scalar_domains,
            node,
            {},
            DynamicRefLimits(),
            None,
            current_sheet="Sheet1",
        )
        assert any(e.kind == "offset-scalar-fallback" for e in events)

    def test_offset_scalar_wide_emits(self) -> None:
        """When _infer_offset_scalar_domains returns >8 values, it emits a wide event."""
        from excel_grapher.core.cell_types import CellKind, CellType, IntervalDomain
        from excel_grapher.core.formula_ast import CellRefNode
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            _infer_offset_scalar_domains,
        )

        env: dict[str, CellType] = {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=1, max=20),
            ),
        }
        node = CellRefNode(address="Sheet1!A1")
        events = self._collect(
            _infer_offset_scalar_domains,
            node,
            env,
            DynamicRefLimits(),
            None,
            current_sheet="Sheet1",
        )
        assert any(e.kind == "offset-scalar-wide" for e in events)


class TestNestedCellTypeTracing:
    """Nested argument-subgraph cells must stream start/done (and cycle) events.

    Issue #111: only top-level infer/expand-env events were emitted, so multiply
    nested formula children in `expand_leaf_env_to_argument_env` produced no
    progress stream.
    """

    def test_nested_formula_chain_emits_start_and_done_per_cell(self) -> None:
        from excel_grapher.core.cell_types import CellKind, CellType, EnumDomain
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            expand_leaf_env_to_argument_env,
            trace_dynamic_refs,
        )

        leaf_env = {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1})),
            ),
        }
        formulas = {
            "Sheet1!B1": "=Sheet1!A1+1",
            "Sheet1!C1": "=Sheet1!B1+1",
            "Sheet1!D1": "=Sheet1!C1+1",
        }

        def get_cell_formula(addr: str) -> str | None:
            return formulas.get(addr)

        def get_refs(formula: str, sheet: str) -> set[str]:
            assert sheet == "Sheet1"
            return {
                formulas["Sheet1!B1"]: {"Sheet1!A1"},
                formulas["Sheet1!C1"]: {"Sheet1!B1"},
                formulas["Sheet1!D1"]: {"Sheet1!C1"},
            }[formula]

        collected: list[DynamicRefTraceEvent] = []
        with trace_dynamic_refs(collected.append):
            expand_leaf_env_to_argument_env(
                {"Sheet1!D1"},
                get_cell_formula,
                get_refs,
                leaf_env,
                DynamicRefLimits(),
            )

        starts = [e for e in collected if e.kind == "cell-type-start"]
        dones = [e for e in collected if e.kind == "cell-type-done"]
        start_addrs = [e.detail["address"] for e in starts]
        done_addrs = [e.detail["address"] for e in dones]

        assert start_addrs == ["Sheet1!D1", "Sheet1!C1", "Sheet1!B1"]
        assert done_addrs == ["Sheet1!B1", "Sheet1!C1", "Sheet1!D1"]
        assert [e.detail["depth"] for e in starts] == [1, 2, 3]
        assert all(e.name == "cell_type_for" for e in starts + dones)
        assert all("formula" in e.detail for e in starts)
        assert all(e.elapsed_s >= 0.0 for e in dones)
        assert all("kind" in e.detail for e in dones)
        # Parent start precedes nested child start; child done precedes parent done.
        assert collected.index(starts[0]) < collected.index(starts[1])
        assert collected.index(dones[0]) < collected.index(dones[1])

    def test_cycle_emits_cell_type_cycle(self) -> None:
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            expand_leaf_env_to_argument_env,
            trace_dynamic_refs,
        )

        formulas = {
            "Sheet1!B1": "=Sheet1!C1",
            "Sheet1!C1": "=Sheet1!B1",
        }

        def get_cell_formula(addr: str) -> str | None:
            return formulas.get(addr)

        def get_refs(formula: str, sheet: str) -> set[str]:
            assert sheet == "Sheet1"
            if "C1" in formula:
                return {"Sheet1!C1"}
            if "B1" in formula:
                return {"Sheet1!B1"}
            return set()

        collected: list[DynamicRefTraceEvent] = []
        with trace_dynamic_refs(collected.append):
            expand_leaf_env_to_argument_env(
                {"Sheet1!B1"},
                get_cell_formula,
                get_refs,
                {},
                DynamicRefLimits(max_depth=4),
            )

        cycles = [e for e in collected if e.kind == "cell-type-cycle"]
        assert len(cycles) >= 1
        assert cycles[0].name == "cell_type_for"
        assert cycles[0].detail["address"] in {"Sheet1!B1", "Sheet1!C1"}
        assert "in_progress" in cycles[0].detail


class TestInferCallTracing:
    """Each OFFSET/INDEX/INDIRECT call should stream start/done, not only a final infer."""

    def test_offset_with_nested_index_emits_per_call_events(self) -> None:
        from excel_grapher.core.cell_types import CellKind, CellType, EnumDomain
        from excel_grapher.grapher.dynamic_refs import (
            infer_dynamic_offset_targets,
            trace_dynamic_refs,
        )

        # OFFSET(INDEX(...), rows, cols) — nested INDEX must appear in the stream
        # before OFFSET completes.
        formula = "=OFFSET(INDEX(Sheet1!A1:A10,Sheet1!B1),Sheet1!C1,0)"
        env = {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2})),
            ),
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0})),
            ),
        }

        collected: list[DynamicRefTraceEvent] = []
        with trace_dynamic_refs(collected.append):
            infer_dynamic_offset_targets(
                formula,
                current_sheet="Sheet1",
                cell_type_env=env,
            )

        call_starts = [e for e in collected if e.kind == "infer-call-start"]
        call_dones = [e for e in collected if e.kind == "infer-call-done"]
        assert any(e.detail.get("function") == "OFFSET" for e in call_starts)
        assert any(e.detail.get("function") == "INDEX" for e in call_starts)
        assert any(e.detail.get("function") == "OFFSET" for e in call_dones)
        assert any(e.detail.get("function") == "INDEX" for e in call_dones)

        offset_start = next(e for e in call_starts if e.detail["function"] == "OFFSET")
        index_start = next(e for e in call_starts if e.detail["function"] == "INDEX")
        index_done = next(e for e in call_dones if e.detail["function"] == "INDEX")
        offset_done = next(e for e in call_dones if e.detail["function"] == "OFFSET")
        assert collected.index(offset_start) < collected.index(index_start)
        assert collected.index(index_start) < collected.index(index_done)
        assert collected.index(index_done) < collected.index(offset_done)

    def test_indirect_emits_infer_call_events(self) -> None:
        from excel_grapher.grapher.dynamic_refs import (
            infer_dynamic_indirect_targets,
            trace_dynamic_refs,
        )

        collected: list[DynamicRefTraceEvent] = []
        with trace_dynamic_refs(collected.append):
            infer_dynamic_indirect_targets(
                '=INDIRECT("Sheet1!A1")',
                current_sheet="Sheet1",
                cell_type_env={},
            )

        starts = [e for e in collected if e.kind == "infer-call-start"]
        dones = [e for e in collected if e.kind == "infer-call-done"]
        assert any(e.detail.get("function") == "INDIRECT" for e in starts)
        assert any(e.detail.get("function") == "INDIRECT" for e in dones)


class TestPersistentCacheHitTracing:
    def test_second_expand_emits_cache_hit(self, tmp_path: Path) -> None:
        from excel_grapher.core.cell_types import CellKind, CellType, EnumDomain
        from excel_grapher.grapher.dynamic_refs import (
            DynamicRefLimits,
            expand_leaf_env_to_argument_env,
            trace_dynamic_refs,
        )
        from excel_grapher.grapher.type_analysis_cache import TypeAnalysisCache

        leaf_env = {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1})),
            ),
        }
        formulas = {"Sheet1!B1": "=Sheet1!A1+1"}

        def get_cell_formula(addr: str) -> str | None:
            return formulas.get(addr)

        def get_refs(formula: str, sheet: str) -> set[str]:
            return {"Sheet1!A1"}

        cache = TypeAnalysisCache.open(tmp_path / "types.sqlite")
        wb_sha = "test_wb_sha"
        try:
            expand_leaf_env_to_argument_env(
                {"Sheet1!B1"},
                get_cell_formula,
                get_refs,
                leaf_env,
                DynamicRefLimits(),
                type_analysis_cache=cache,
                workbook_sha256=wb_sha,
            )
            cache.flush()

            collected: list[DynamicRefTraceEvent] = []
            with trace_dynamic_refs(collected.append):
                expand_leaf_env_to_argument_env(
                    {"Sheet1!B1"},
                    get_cell_formula,
                    get_refs,
                    leaf_env,
                    DynamicRefLimits(),
                    type_analysis_cache=cache,
                    workbook_sha256=wb_sha,
                )

            hits = [e for e in collected if e.kind == "cell-type-cache-hit"]
            assert len(hits) == 1
            assert hits[0].detail["address"] == "Sheet1!B1"
            assert hits[0].detail["source"] == "persistent"
            assert not any(e.kind == "cell-type-start" for e in collected)
        finally:
            cache.close()


class TestBuilderDynamicRefCellTracing:
    def test_graph_build_streams_cell_and_nested_children(self, tmp_path: Path) -> None:
        from excel_grapher.core.cell_types import CellKind, CellType, EnumDomain
        from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefLimits

        excel_path = tmp_path / "nested_offset.xlsx"
        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        ws.write_number(0, 0, 0)  # A1 base
        ws.write_number(0, 11, 1)  # L1 leaf
        ws.write_formula(0, 12, "=Sheet1!L1", None, 1)  # M1 intermediate
        ws.write_formula(0, 2, "=OFFSET(Sheet1!A1,Sheet1!M1,0)", None, 0)  # C1
        wb.close()

        config = DynamicRefConfig(
            cell_type_env={
                "Sheet1!L1": CellType(
                    kind=CellKind.NUMBER,
                    enum=EnumDomain(values=frozenset({0, 1})),
                ),
            },
            limits=DynamicRefLimits(),
        )

        collected: list[DynamicRefTraceEvent] = []
        with trace_dynamic_refs(collected.append):
            create_dependency_graph(
                excel_path,
                ["Sheet1!C1"],
                load_values=False,
                dynamic_refs=config,
            )

        cell_starts = [e for e in collected if e.kind == "dynamic-ref-cell-start"]
        cell_dones = [e for e in collected if e.kind == "dynamic-ref-cell-done"]
        assert any(e.detail.get("address") == "Sheet1!C1" for e in cell_starts)
        assert any(e.detail.get("address") == "Sheet1!C1" for e in cell_dones)

        nested_starts = [e for e in collected if e.kind == "cell-type-start"]
        assert any(e.detail.get("address") == "Sheet1!M1" for e in nested_starts)

        infer_calls = [e for e in collected if e.kind == "infer-call-start"]
        assert any(e.detail.get("function") == "OFFSET" for e in infer_calls)

        # Nested child work is enclosed by the parent cell start/done.
        cell_start = next(e for e in cell_starts if e.detail["address"] == "Sheet1!C1")
        cell_done = next(e for e in cell_dones if e.detail["address"] == "Sheet1!C1")
        nested = next(e for e in nested_starts if e.detail["address"] == "Sheet1!M1")
        assert collected.index(cell_start) < collected.index(nested) < collected.index(cell_done)
