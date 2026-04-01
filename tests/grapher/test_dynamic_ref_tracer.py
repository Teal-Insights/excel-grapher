"""Tests for the first-class dynamic-ref tracing infrastructure."""

from __future__ import annotations

import pytest

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
            event.kind = "other"  # type: ignore[misc]  # ty: ignore[invalid-assignment]

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
        assert any(e.kind == "expand-env" and e.name == "expand_leaf_env_to_argument_env" for e in events)

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
        assert any(e.kind == "expand-env-error" and e.name == "expand_leaf_env_to_argument_env" for e in events)

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
