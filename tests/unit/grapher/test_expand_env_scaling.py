"""Scaling guards for `expand_leaf_env_to_argument_env` (issue #463).

Covers three defences against the LIC-DSF hang:

- Consumed-leaf bookkeeping is only maintained when a persistent
  `TypeAnalysisCache` is actually going to read it.
- `expand-env-progress` traces so callers are not blind during a long expansion.
- A depth guard that fails with `DynamicRefError` instead of `RecursionError`.
"""

from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path

import pytest

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
)
from excel_grapher.grapher import dynamic_refs as dynamic_refs_mod
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefError,
    DynamicRefLimits,
    DynamicRefTraceEvent,
    expand_leaf_env_to_argument_env,
    trace_dynamic_refs,
)
from excel_grapher.grapher.type_analysis_cache import TypeAnalysisCache

_ChainCallables = tuple[Callable[[str], str | None], Callable[[str, str], set[str]], CellTypeEnv]


def _build_chain(depth: int) -> _ChainCallables:
    """Build an in-memory chain `A1` (leaf) → `A2 = A1+1` → ... → `A{depth}`.

    Returns the `get_cell_formula` / `get_refs_from_formula` callables and the
    leaf env, ready to pass to `expand_leaf_env_to_argument_env`.
    """
    formulas = {f"Sheet1!A{r}": f"=Sheet1!A{r - 1}+1" for r in range(2, depth + 1)}
    refs = {f"=Sheet1!A{r - 1}+1": {f"Sheet1!A{r - 1}"} for r in range(2, depth + 1)}

    def get_cell_formula(addr: str) -> str | None:
        return formulas.get(addr)

    def get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        return set(refs.get(formula, set()))

    leaf_env: CellTypeEnv = {
        "Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2})))
    }
    return get_cell_formula, get_refs_from_formula, leaf_env


def _expand_with_trace(
    depth: int, **kwargs: object
) -> tuple[list[DynamicRefTraceEvent], dict[str, CellType]]:
    """Run an expansion over a chain of `depth`, collecting trace events."""
    get_cell_formula, get_refs_from_formula, leaf_env = _build_chain(depth)
    events: list[DynamicRefTraceEvent] = []
    with trace_dynamic_refs(events.append):
        env = expand_leaf_env_to_argument_env(
            {f"Sheet1!A{depth}"},
            get_cell_formula,
            get_refs_from_formula,
            leaf_env,
            DynamicRefLimits(),
            **kwargs,
        )
    return events, env


class TestConsumedLeafGating:
    """`_consumed_leaves` is only read by `_persist_result`, so skip it otherwise."""

    def test_tracking_disabled_without_persistent_cache(self) -> None:
        events, env = _expand_with_trace(6)

        assert env["Sheet1!A6"].enum is not None
        (expand,) = [e for e in events if e.kind == "expand-env"]
        assert expand.detail["consumed_leaf_tracking"] is False

    def test_tracking_enabled_with_persistent_cache(self, tmp_path: Path) -> None:
        cache = TypeAnalysisCache.open(tmp_path / "tac.sqlite3")
        try:
            events, env = _expand_with_trace(
                6,
                type_analysis_cache=cache,
                workbook_sha256="wb_463",
            )
            cache.flush()
        finally:
            cache.close()

        assert env["Sheet1!A6"].enum is not None
        (expand,) = [e for e in events if e.kind == "expand-env"]
        assert expand.detail["consumed_leaf_tracking"] is True


class TestExpandProgressTraces:
    def test_progress_events_emitted_for_long_expansion(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(dynamic_refs_mod, "_EXPAND_PROGRESS_INTERVAL", 5)
        events, _ = _expand_with_trace(40)

        progress = [e for e in events if e.kind == "expand-env-progress"]
        assert progress, "expected expand-env-progress events during a long expansion"
        assert all(e.name == "expand_leaf_env_to_argument_env" for e in progress)
        assert all("calls" in e.detail and "cache_size" in e.detail for e in progress)
        # Progress must be reported before the terminal expand-env event.
        assert events.index(progress[-1]) < events.index(
            next(e for e in events if e.kind == "expand-env")
        )

    def test_no_progress_events_for_short_expansion(self) -> None:
        events, _ = _expand_with_trace(3)
        assert [e for e in events if e.kind == "expand-env-progress"] == []


class TestAnalysisDepthGuard:
    def test_deep_chain_raises_dynamic_ref_error(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(dynamic_refs_mod, "_MAX_ANALYSIS_DEPTH", 20)

        with pytest.raises(DynamicRefError) as excinfo:
            _expand_with_trace(60)

        message = str(excinfo.value)
        assert "depth" in message.lower()
        assert "Sheet1!A" in message

    def test_default_depth_limit_leaves_recursion_headroom(self) -> None:
        """The guard must trip before CPython's own recursion limit does."""
        assert sys.getrecursionlimit() * 0.75 > dynamic_refs_mod._MAX_ANALYSIS_DEPTH

    def test_chain_within_limit_still_succeeds(self) -> None:
        _, env = _expand_with_trace(50)
        assert env["Sheet1!A50"].enum is not None
