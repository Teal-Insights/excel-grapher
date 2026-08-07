"""Scaling guards for `expand_leaf_env_to_argument_env` (issues #463, #465).

Covers four defences against the LIC-DSF hang:

- Consumed-leaf bookkeeping is only maintained when a persistent
  `TypeAnalysisCache` is actually going to read it.
- `expand-env-progress` traces so callers are not blind during a long expansion.
- A depth guard that fails with `DynamicRefError` instead of `RecursionError`.
- Already-inferred refs are served in bulk, so a long chain over a large static
  range costs `O(depth + range_cells)` `cell_type_for` entries, not
  `O(depth x range_cells)`.
- The address expansion of a static range is memoized, so re-mentioning the same
  range at every level of a chain does not re-derive its cells every time.
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
    normalize_cell_type_env_key,
)
from excel_grapher.core.formula_ast import RangeNode
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


def _build_range_chain(depth: int, width: int, *, cell_value: int = 0) -> _ChainCallables:
    """Build a chain whose every level re-mentions the same static range.

    `Sheet1!A{r} = Sheet1!A{r-1} + SUM(Sheet1!C1:C{width})`, with `Sheet1!A1` and
    the whole `C` range as leaves.  This is the shape from issue #465: the AST
    collector expands the range into `width` addresses at every one of the
    `depth` levels.
    """
    rng = f"Sheet1!C1:C{width}"
    formulas = {f"Sheet1!A{r}": f"=Sheet1!A{r - 1}+SUM({rng})" for r in range(2, depth + 1)}

    def get_cell_formula(addr: str) -> str | None:
        return formulas.get(addr)

    def get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        # Static addresses are contributed by `_collect_static_addresses_from_ast`.
        return set()

    leaf_env: CellTypeEnv = {
        "Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))
    }
    for row in range(1, width + 1):
        leaf_env[f"Sheet1!C{row}"] = CellType(
            kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({cell_value}))
        )
    return get_cell_formula, get_refs_from_formula, leaf_env


def _expand_range_chain_with_trace(
    depth: int, width: int, *, cell_value: int = 0, **kwargs: object
) -> tuple[list[DynamicRefTraceEvent], dict[str, CellType]]:
    """Run an expansion over `_build_range_chain`, collecting trace events."""
    get_cell_formula, get_refs_from_formula, leaf_env = _build_range_chain(
        depth, width, cell_value=cell_value
    )
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


class TestBulkCachedRefResolution:
    """Refs already in the in-memory cache are served without re-entering `cell_type_for`."""

    def test_call_volume_is_not_depth_times_range(self) -> None:
        depth, width = 40, 100
        events, env = _expand_range_chain_with_trace(depth, width)

        assert env[f"Sheet1!A{depth}"].enum is not None
        (expand,) = [e for e in events if e.kind == "expand-env"]
        calls = expand.detail["calls"]
        assert isinstance(calls, int)
        # O(depth + range_cells), not O(depth x range_cells) (= 4000 here).
        assert calls <= 3 * depth + width + 25, f"expected additive call volume, got {calls}"

    def test_call_volume_grows_additively_with_depth(self) -> None:
        width = 100
        (shallow,) = [
            e for e in _expand_range_chain_with_trace(10, width)[0] if e.kind == "expand-env"
        ]
        (deep,) = [
            e for e in _expand_range_chain_with_trace(40, width)[0] if e.kind == "expand-env"
        ]

        shallow_calls = shallow.detail["calls"]
        deep_calls = deep.detail["calls"]
        assert isinstance(shallow_calls, int)
        assert isinstance(deep_calls, int)
        # Quadrupling the depth must not multiply the call volume by the range width.
        assert deep_calls - shallow_calls <= 3 * (40 - 10) + 10

    def test_bulk_served_refs_are_reported(self) -> None:
        events, _ = _expand_range_chain_with_trace(20, 50)

        (expand,) = [e for e in events if e.kind == "expand-env"]
        bulk = expand.detail["bulk_ref_hits"]
        assert isinstance(bulk, int)
        # Every level past the first re-mentions the whole range.
        assert bulk >= 50 * 18

    def test_inferred_types_are_unchanged_by_bulk_serving(self) -> None:
        depth, width, value = 6, 4, 3
        _, env = _expand_range_chain_with_trace(depth, width, cell_value=value)

        # A1 = 1, and each of the `depth - 1` levels adds SUM(C1:C{width}).
        expected = 1 + (depth - 1) * width * value
        top = env[f"Sheet1!A{depth}"].enum
        assert top is not None
        assert top.values == frozenset({expected})

    def test_progress_traces_still_fire_when_refs_are_bulk_served(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Bulk-served refs still count toward progress, so long runs stay visible."""
        monkeypatch.setattr(dynamic_refs_mod, "_EXPAND_PROGRESS_INTERVAL", 100)
        events, _ = _expand_range_chain_with_trace(30, 60)

        progress = [e for e in events if e.kind == "expand-env-progress"]
        assert progress, "expected progress traces while bulk-serving cached refs"

    def test_persistent_cache_still_invalidates_on_range_leaf_change(self, tmp_path: Path) -> None:
        """Bulk serving must not drop consumed-leaf bookkeeping for cached refs."""
        cache = TypeAnalysisCache.open(tmp_path / "tac.sqlite3")
        try:
            _, env_zero = _expand_range_chain_with_trace(
                5, 4, cell_value=0, type_analysis_cache=cache, workbook_sha256="wb_465"
            )
            cache.flush()
            _, env_three = _expand_range_chain_with_trace(
                5, 4, cell_value=3, type_analysis_cache=cache, workbook_sha256="wb_465"
            )
            cache.flush()
        finally:
            cache.close()

        zero_enum = env_zero["Sheet1!A5"].enum
        three_enum = env_three["Sheet1!A5"].enum
        assert zero_enum is not None
        assert three_enum is not None
        assert zero_enum.values == frozenset({1})
        assert three_enum.values == frozenset({1 + 4 * 4 * 3})


class TestStaticRangeExpansionMemo:
    """Expanding the same static range at every level should not redo the work."""

    def test_chain_expansion_reuses_memoized_range_addresses(self) -> None:
        dynamic_refs_mod._expanded_range_keys.cache_clear()
        dynamic_refs_mod._range_node_cell_addresses.cache_clear()

        depth, width = 20, 50
        _expand_range_chain_with_trace(depth, width)

        # One miss for the range, a hit for every level that re-mentions it.
        collect_info = dynamic_refs_mod._expanded_range_keys.cache_info()
        assert collect_info.hits >= depth - 2
        infer_info = dynamic_refs_mod._range_node_cell_addresses.cache_info()
        assert infer_info.hits >= depth - 2

    def test_memoized_expansion_matches_direct_expansion(self) -> None:
        keys = dynamic_refs_mod._expanded_range_keys(
            sheet="Sheet1",
            start_col="B",
            start_row=2,
            end_col="C",
            end_row=3,
            max_cells=5000,
        )
        assert keys == ("Sheet1!B2", "Sheet1!C2", "Sheet1!B3", "Sheet1!C3")

    def test_memo_respects_the_max_cells_limit(self) -> None:
        """`max_cells` is part of the key, so truncation is not cached across limits."""
        args = {
            "sheet": "Sheet1",
            "start_col": "A",
            "start_row": 1,
            "end_col": "A",
            "end_row": 10,
        }
        assert len(dynamic_refs_mod._expanded_range_keys(**args, max_cells=5000)) == 10
        # Over the limit, `expand_range` degrades to the two endpoints.
        assert dynamic_refs_mod._expanded_range_keys(**args, max_cells=4) == (
            "Sheet1!A1",
            "Sheet1!A10",
        )

    def test_range_node_addresses_are_row_major(self) -> None:
        node = RangeNode(start="Sheet1!B2", end="Sheet1!C3")
        assert dynamic_refs_mod._range_node_cell_addresses(node) == (
            "Sheet1!B2",
            "Sheet1!C2",
            "Sheet1!B3",
            "Sheet1!C3",
        )

    def test_range_node_addresses_are_already_env_normalized(self) -> None:
        """SUM inference looks these up in the env directly, so they must be canonical."""
        node = RangeNode(start="'My Sheet'!$b$2", end="'My Sheet'!$C$3")
        addrs = dynamic_refs_mod._range_node_cell_addresses(node)

        assert addrs is not None
        assert addrs[0] == "My Sheet!B2"
        assert all(addr == normalize_cell_type_env_key(addr) for addr in addrs)


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
