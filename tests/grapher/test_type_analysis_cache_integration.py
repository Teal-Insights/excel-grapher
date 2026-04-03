"""Integration tests: persistent cache inside expand_leaf_env_to_argument_env (Phase C & D)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
)
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefLimits,
    expand_leaf_env_to_argument_env,
)
from excel_grapher.grapher.type_analysis_cache import (
    TypeAnalysisCache,
    _compute_limits_fingerprint,
)


def _make_env(mapping: dict[str, CellType]) -> CellTypeEnv:
    return mapping


# ---------------------------------------------------------------------------
# Helpers to build a small formula graph for testing
# ---------------------------------------------------------------------------

# Formula graph:
#   Sheet1!C1 = Sheet1!A1 + Sheet1!B1  (intermediate formula)
#   Sheet1!A1 = leaf (constrained)
#   Sheet1!B1 = leaf (constrained)
#
# We want to analyse Sheet1!C1's type.

_FORMULAS = {
    "Sheet1!C1": "=Sheet1!A1+Sheet1!B1",
}

_REFS = {
    "=Sheet1!A1+Sheet1!B1": {"Sheet1!A1", "Sheet1!B1"},
}


def _get_cell_formula(addr: str) -> str | None:
    return _FORMULAS.get(addr)


def _get_refs(formula: str, sheet: str) -> set[str]:
    return _REFS.get(formula, set())


_LEAF_ENV: CellTypeEnv = _make_env(
    {
        "Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2}))),
        "Sheet1!B1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({10, 20}))),
    }
)

_LIMITS = DynamicRefLimits()
_WB_SHA = "fake_workbook_sha256"
_LIMITS_FP = _compute_limits_fingerprint(_LIMITS)


# ---------------------------------------------------------------------------
# Phase C – Integration tests
# ---------------------------------------------------------------------------


class TestPersistentCacheIntegration:
    def test_second_run_reuses_cached_type(self, tmp_path: Path) -> None:
        """Second run should hit the persistent cache and not recompute."""
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            # First run: computes C1
            env1 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            cache.flush()
            assert "Sheet1!C1" in env1
            assert env1["Sheet1!C1"].kind is CellKind.NUMBER

            # Second run: should reuse from persistent cache
            # Use a fresh in-memory cache to prove it loads from SQLite
            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            assert env2["Sheet1!C1"] == env1["Sheet1!C1"]
            assert cache.stats.hits >= 1
        finally:
            cache.close()

    def test_leaf_cells_prefer_explicit_constraints(self, tmp_path: Path) -> None:
        """leaf_env entries always dominate over cached formula-cell output."""
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env = expand_leaf_env_to_argument_env(
                {"Sheet1!A1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            # A1 is a leaf; its type comes from leaf_env, not the cache
            assert env["Sheet1!A1"] == _LEAF_ENV["Sheet1!A1"]
        finally:
            cache.close()

    def test_constraint_change_invalidates_cached_entry(self, tmp_path: Path) -> None:
        """Changing a consumed leaf constraint should invalidate the cache."""
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            # First run with original constraints
            expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            cache.flush()

            # Second run with different leaf constraints
            new_leaf_env = _make_env(
                {
                    "Sheet1!A1": CellType(
                        kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({100}))
                    ),
                    "Sheet1!B1": CellType(
                        kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({200}))
                    ),
                }
            )
            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                new_leaf_env,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            # C1 should be recomputed with new constraints
            assert env2["Sheet1!C1"].enum is not None
            assert 300 in env2["Sheet1!C1"].enum.values
        finally:
            cache.close()

    def test_workbook_edit_invalidates_old_entry(self, tmp_path: Path) -> None:
        """A different workbook hash should not hit the old cache entry."""
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env1 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256="wb_v1",
            )
            cache.flush()
            hits_before = cache.stats.hits

            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256="wb_v2",
            )
            # Should not have gotten a cache hit for wb_v2
            assert cache.stats.hits == hits_before
            # But results should still be correct
            assert env2["Sheet1!C1"] == env1["Sheet1!C1"]
        finally:
            cache.close()

    def test_unrelated_constraint_change_does_not_invalidate(self, tmp_path: Path) -> None:
        """A leaf constraint outside the consumed set should not bust the entry."""
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env1 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            cache.flush()
            hits_before = cache.stats.hits

            # Add an unrelated leaf constraint
            extended_env = dict(_LEAF_ENV)
            extended_env["Sheet1!Z99"] = CellType(
                kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"irrelevant"}))
            )
            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                extended_env,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            assert cache.stats.hits > hits_before
            assert env2["Sheet1!C1"] == env1["Sheet1!C1"]
        finally:
            cache.close()

    def test_formula_change_invalidates_old_entry(self, tmp_path: Path) -> None:
        """Same address with a changed formula should miss."""
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            cache.flush()
            hits_before = cache.stats.hits

            # Different formula for C1
            def alt_formula(addr: str) -> str | None:
                if addr == "Sheet1!C1":
                    return "=Sheet1!A1*Sheet1!B1"
                return None

            alt_refs = {"=Sheet1!A1*Sheet1!B1": {"Sheet1!A1", "Sheet1!B1"}}

            def alt_get_refs(formula: str, sheet: str) -> set[str]:
                return alt_refs.get(formula, set())

            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                alt_formula,
                alt_get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            # Formula changed, so no cache hit
            assert cache.stats.hits == hits_before
            # Result should reflect multiplication, not addition
            assert env2["Sheet1!C1"].enum is not None
            assert 20 in env2["Sheet1!C1"].enum.values  # 1*20 or 2*10
        finally:
            cache.close()

    def test_version_bump_invalidates_old_entry(self, tmp_path: Path) -> None:
        """Old rows should not be trusted after version change."""
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                _get_cell_formula,
                _get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            cache.flush()
            hits_before = cache.stats.hits

            # Simulate a version bump
            import excel_grapher.grapher.type_analysis_cache as tac_mod

            old_version = tac_mod._EXCEL_GRAPHER_VERSION
            try:
                tac_mod._EXCEL_GRAPHER_VERSION = "99.99.99"
                expand_leaf_env_to_argument_env(
                    {"Sheet1!C1"},
                    _get_cell_formula,
                    _get_refs,
                    _LEAF_ENV,
                    _LIMITS,
                    type_analysis_cache=cache,
                    workbook_sha256=_WB_SHA,
                )
                assert cache.stats.hits == hits_before
            finally:
                tac_mod._EXCEL_GRAPHER_VERSION = old_version
        finally:
            cache.close()

    def test_recursive_chain_hydrates_from_cache(self, tmp_path: Path) -> None:
        """Multi-step formula chain should be reused on a second run."""
        # D1 = C1 + 1, C1 = A1 + B1
        formulas = {
            "Sheet1!D1": "=Sheet1!C1+1",
            "Sheet1!C1": "=Sheet1!A1+Sheet1!B1",
        }
        refs = {
            "=Sheet1!C1+1": {"Sheet1!C1"},
            "=Sheet1!A1+Sheet1!B1": {"Sheet1!A1", "Sheet1!B1"},
        }

        def get_formula(addr: str) -> str | None:
            return formulas.get(addr)

        def get_refs(formula: str, sheet: str) -> set[str]:
            return refs.get(formula, set())

        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env1 = expand_leaf_env_to_argument_env(
                {"Sheet1!D1"},
                get_formula,
                get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            cache.flush()
            assert "Sheet1!D1" in env1
            assert "Sheet1!C1" in env1

            # Second run — should use cache for D1 (and C1 if traversed)
            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!D1"},
                get_formula,
                get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            assert env2["Sheet1!D1"] == env1["Sheet1!D1"]
            # D1 loaded from persistent cache so no traversal of C1 needed
            assert cache.stats.hits >= 1
        finally:
            cache.close()


class TestConsumedLeafRestorationOnCacheHit:
    """Fix 2: When a persistent cache hit occurs for a middle cell, its
    consumed-leaf keys must be restored so that ancestors record the
    transitive leaf dependencies.  Otherwise, changing the bottom leaf
    won't invalidate the top cell on a subsequent run.
    """

    # Three-level chain:
    #   E1 = D1 + B1
    #   D1 = B1 + C1
    #   B1 = A1 + 1
    #   C1 = A1 * 2
    #   A1 = leaf

    _FORMULAS: dict[str, str | None] = {
        "Sheet1!E1": "=Sheet1!D1+Sheet1!B1",
        "Sheet1!D1": "=Sheet1!B1+Sheet1!C1",
        "Sheet1!B1": "=Sheet1!A1+1",
        "Sheet1!C1": "=Sheet1!A1*2",
    }
    _REFS: dict[str, set[str]] = {
        "=Sheet1!D1+Sheet1!B1": {"Sheet1!D1", "Sheet1!B1"},
        "=Sheet1!B1+Sheet1!C1": {"Sheet1!B1", "Sheet1!C1"},
        "=Sheet1!A1+1": {"Sheet1!A1"},
        "=Sheet1!A1*2": {"Sheet1!A1"},
    }

    def _get_formula(self, addr: str) -> str | None:
        return self._FORMULAS.get(addr)

    def _get_refs(self, formula: str, sheet: str) -> set[str]:
        return self._REFS.get(formula, set())

    def test_middle_cache_hit_propagates_leaves_to_top(self, tmp_path: Path) -> None:
        """Changing A1 must invalidate E1 even when B1 is served from persistent cache."""
        leaf_env_v1: CellTypeEnv = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER,
                    enum=EnumDomain(values=frozenset({1, 2})),
                ),
            }
        )

        # Run 1: populate the persistent cache for B1, C1, D1
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env1 = expand_leaf_env_to_argument_env(
                {"Sheet1!D1"},
                self._get_formula,
                self._get_refs,
                leaf_env_v1,
                DynamicRefLimits(),
                type_analysis_cache=cache,
                workbook_sha256="wb_fix2",
            )
            cache.flush()
            assert env1["Sheet1!D1"].enum is not None
        finally:
            cache.close()

        # Run 2: analyse E1 (which references D1 and B1).
        # B1 may be served from persistent cache.
        cache2 = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!E1"},
                self._get_formula,
                self._get_refs,
                leaf_env_v1,
                DynamicRefLimits(),
                type_analysis_cache=cache2,
                workbook_sha256="wb_fix2",
            )
            cache2.flush()
            assert env2["Sheet1!E1"].enum is not None
        finally:
            cache2.close()

        # Run 3: change A1's constraint.  E1 must be invalidated.
        leaf_env_v2: CellTypeEnv = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER,
                    enum=EnumDomain(values=frozenset({100, 200})),
                ),
            }
        )

        cache3 = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env3 = expand_leaf_env_to_argument_env(
                {"Sheet1!E1"},
                self._get_formula,
                self._get_refs,
                leaf_env_v2,
                DynamicRefLimits(),
                type_analysis_cache=cache3,
                workbook_sha256="wb_fix2",
            )
            # E1 must reflect the new A1 values, not stale cached results
            assert env3["Sheet1!E1"].enum is not None
            # With A1∈{100,200}: B1∈{101,201}, C1∈{200,400},
            # D1=B1+C1∈{301,401,501,601}, E1=D1+B1
            # All values should be ≥ 402
            assert all(v >= 402 for v in env3["Sheet1!E1"].enum.values), (
                f"E1 should reflect updated A1 but got {env3['Sheet1!E1'].enum.values}"
            )
        finally:
            cache3.close()


# ---------------------------------------------------------------------------
# Phase D – Crash-safe partial progress
# ---------------------------------------------------------------------------


class TestPartialProgress:
    def test_partial_run_leaves_reusable_entries(self, tmp_path: Path) -> None:
        """Completed cells survive even if analysis stops early."""
        # C1 = A1+B1 (will succeed), D1 depends on missing leaf (will fail)
        formulas = {
            "Sheet1!C1": "=Sheet1!A1+Sheet1!B1",
            "Sheet1!D1": "=Sheet1!X1+1",
        }
        refs = {
            "=Sheet1!A1+Sheet1!B1": {"Sheet1!A1", "Sheet1!B1"},
            "=Sheet1!X1+1": {"Sheet1!X1"},
        }

        def get_formula(addr: str) -> str | None:
            return formulas.get(addr)

        def get_refs(formula: str, sheet: str) -> set[str]:
            return refs.get(formula, set())

        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            # This should raise because X1 has no formula and no leaf constraint
            from excel_grapher.grapher.dynamic_refs import DynamicRefError

            with pytest.raises(DynamicRefError):
                expand_leaf_env_to_argument_env(
                    {"Sheet1!C1", "Sheet1!D1"},
                    get_formula,
                    get_refs,
                    _LEAF_ENV,
                    _LIMITS,
                    type_analysis_cache=cache,
                    workbook_sha256=_WB_SHA,
                )
            cache.flush()
        finally:
            cache.close()

        # Reopen and check C1 was persisted (it was analyzed before D1 failed)
        cache2 = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!C1"},
                get_formula,
                get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache2,
                workbook_sha256=_WB_SHA,
            )
            assert cache2.stats.hits >= 1
            assert env2["Sheet1!C1"].kind is CellKind.NUMBER
        finally:
            cache2.close()

    def test_best_effort_under_mixed_success(self, tmp_path: Path) -> None:
        """Successful inferred cells are persisted even when others fail."""
        # C1 succeeds; analysis of D1 = unsupported FOO() → gets ANY
        formulas = {
            "Sheet1!C1": "=Sheet1!A1+Sheet1!B1",
            "Sheet1!D1": "=FOO(Sheet1!A1)",
        }
        refs = {
            "=Sheet1!A1+Sheet1!B1": {"Sheet1!A1", "Sheet1!B1"},
            "=FOO(Sheet1!A1)": {"Sheet1!A1"},
        }

        def get_formula(addr: str) -> str | None:
            return formulas.get(addr)

        def get_refs(formula: str, sheet: str) -> set[str]:
            return refs.get(formula, set())

        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env = expand_leaf_env_to_argument_env(
                {"Sheet1!C1", "Sheet1!D1"},
                get_formula,
                get_refs,
                _LEAF_ENV,
                _LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=_WB_SHA,
            )
            cache.flush()
            # C1 should be NUMBER with enum, D1 should be ANY
            assert env["Sheet1!C1"].kind is CellKind.NUMBER
            assert env["Sheet1!D1"].kind is CellKind.ANY
            # Only C1 should be cached (not ANY results per design)
            assert cache.stats.writes >= 1
        finally:
            cache.close()


# ---------------------------------------------------------------------------
# Phase E – Transitive consumed-leaf tracking
# ---------------------------------------------------------------------------


class TestTransitiveConsumedLeaves:
    """Verify that changing a leaf constraint invalidates cached results for
    *all* ancestor formula cells, not just the immediate parent.

    Diamond graph:
        D1 = B1 + C1   (grandparent)
        B1 = A1 + 1    (parent, left)
        C1 = A1 * 2    (parent, right)
        A1 = leaf       (constrained)

    If A1's constraint changes, D1's cached entry must be invalidated even
    though D1 does not reference A1 directly.  This requires consumed-leaf
    tracking to propagate transitively from children to parents on the
    analysis stack.
    """

    _FORMULAS: dict[str, str | None] = {
        "Sheet1!D1": "=Sheet1!B1+Sheet1!C1",
        "Sheet1!B1": "=Sheet1!A1+1",
        "Sheet1!C1": "=Sheet1!A1*2",
    }
    _REFS: dict[str, set[str]] = {
        "=Sheet1!B1+Sheet1!C1": {"Sheet1!B1", "Sheet1!C1"},
        "=Sheet1!A1+1": {"Sheet1!A1"},
        "=Sheet1!A1*2": {"Sheet1!A1"},
    }
    _LIMITS = DynamicRefLimits()
    _WB_SHA = "transitive_test_wb_sha"

    def _get_formula(self, addr: str) -> str | None:
        return self._FORMULAS.get(addr)

    def _get_refs(self, formula: str, sheet: str) -> set[str]:
        return self._REFS.get(formula, set())

    def test_upstream_leaf_change_invalidates_grandparent(self, tmp_path: Path) -> None:
        """D1's cached type must not be reused when A1's constraint changes."""
        leaf_env_v1: CellTypeEnv = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER,
                    enum=EnumDomain(values=frozenset({1, 2})),
                ),
            }
        )

        # Run 1: populate the persistent cache
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env1 = expand_leaf_env_to_argument_env(
                {"Sheet1!D1"},
                self._get_formula,
                self._get_refs,
                leaf_env_v1,
                self._LIMITS,
                type_analysis_cache=cache,
                workbook_sha256=self._WB_SHA,
            )
            cache.flush()
            # D1 = B1 + C1; B1 ∈ {2,3}, C1 ∈ {2,4} → D1 ∈ {4,5,6,7}
            assert env1["Sheet1!D1"].enum is not None
            assert env1["Sheet1!D1"].enum.values == frozenset({4, 5, 6, 7})
        finally:
            cache.close()

        # Run 2: change A1's constraint and reopen the cache
        leaf_env_v2: CellTypeEnv = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER,
                    enum=EnumDomain(values=frozenset({10, 20})),
                ),
            }
        )

        cache2 = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            env2 = expand_leaf_env_to_argument_env(
                {"Sheet1!D1"},
                self._get_formula,
                self._get_refs,
                leaf_env_v2,
                self._LIMITS,
                type_analysis_cache=cache2,
                workbook_sha256=self._WB_SHA,
            )
            # D1 = B1 + C1; B1 ∈ {11,21}, C1 ∈ {20,40} → D1 ∈ {31,41,51,61}
            assert env2["Sheet1!D1"].enum is not None
            assert env2["Sheet1!D1"].enum.values == frozenset({31, 41, 51, 61}), (
                f"D1 should reflect updated A1 constraints but got "
                f"{env2['Sheet1!D1'].enum.values}; "
                f"stale cache entry was likely reused"
            )
        finally:
            cache2.close()
