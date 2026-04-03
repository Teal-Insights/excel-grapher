"""Tests for persistent type-analysis cache (Phase A & B)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    EnumDomain,
    GreaterThanCell,
    IntervalDomain,
    NotEqualCell,
    RealIntervalDomain,
)
from excel_grapher.grapher.dynamic_refs import DynamicRefLimits
from excel_grapher.grapher.type_analysis_cache import (
    TypeAnalysisCache,
    _cell_type_from_json,
    _cell_type_to_json,
    _compute_key_hash,
    _compute_leaf_env_subset_fingerprint,
    _compute_limits_fingerprint,
)

# ---------------------------------------------------------------------------
# Phase A – Round-trip CellType serialization
# ---------------------------------------------------------------------------


class TestCellTypeSerialization:
    def test_round_trip_number_enum(self) -> None:
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2, 3})))
        assert _cell_type_from_json(_cell_type_to_json(ct)) == ct

    def test_round_trip_number_interval(self) -> None:
        ct = CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=100))
        assert _cell_type_from_json(_cell_type_to_json(ct)) == ct

    def test_round_trip_string_enum(self) -> None:
        ct = CellType(kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"a", "b"})))
        assert _cell_type_from_json(_cell_type_to_json(ct)) == ct

    def test_round_trip_bool_enum(self) -> None:
        ct = CellType(kind=CellKind.BOOL, enum=EnumDomain(values=frozenset({True, False})))
        assert _cell_type_from_json(_cell_type_to_json(ct)) == ct

    def test_round_trip_any(self) -> None:
        ct = CellType(kind=CellKind.ANY)
        assert _cell_type_from_json(_cell_type_to_json(ct)) == ct

    def test_round_trip_real_interval(self) -> None:
        ct = CellType(kind=CellKind.NUMBER, real_interval=RealIntervalDomain(min=0.5, max=9.9))
        assert _cell_type_from_json(_cell_type_to_json(ct)) == ct

    def test_round_trip_mixed_numeric_enum(self) -> None:
        """int vs float distinction must be preserved."""
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2.5})))
        result = _cell_type_from_json(_cell_type_to_json(ct))
        assert result == ct
        assert result.enum is not None
        vals = result.enum.values
        assert 1 in vals and isinstance(next(v for v in vals if v == 1), int)
        assert 2.5 in vals and isinstance(next(v for v in vals if v == 2.5), float)

    def test_round_trip_preserves_bool_type(self) -> None:
        """Booleans must round-trip as bool, not int."""
        ct = CellType(kind=CellKind.BOOL, enum=EnumDomain(values=frozenset({True})))
        result = _cell_type_from_json(_cell_type_to_json(ct))
        assert result == ct
        assert result.enum is not None
        val = next(iter(result.enum.values))
        assert isinstance(val, bool)

    def test_round_trip_relations(self) -> None:
        """CellType with relations must survive a JSON round-trip."""
        ct = CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=0, max=100),
            relations=(NotEqualCell("Sheet1!Y1"), GreaterThanCell("Sheet1!X1")),
        )
        result = _cell_type_from_json(_cell_type_to_json(ct))
        assert result.kind == ct.kind
        assert result.interval == ct.interval
        assert result.enum == ct.enum
        assert result.real_interval == ct.real_interval
        assert set(result.relations) == set(ct.relations)

    def test_relations_difference_changes_json(self) -> None:
        """Two CellTypes differing only in relations must produce different JSON."""
        ct1 = CellType(
            kind=CellKind.NUMBER,
            relations=(NotEqualCell("Sheet1!A1"),),
        )
        ct2 = CellType(
            kind=CellKind.NUMBER,
            relations=(GreaterThanCell("Sheet1!A1"),),
        )
        assert _cell_type_to_json(ct1) != _cell_type_to_json(ct2)


# ---------------------------------------------------------------------------
# Phase A – Cache miss / hit behavior
# ---------------------------------------------------------------------------


class TestCacheLookup:
    def test_miss_on_empty_cache(self, tmp_path: Path) -> None:
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            result = cache.get_formula_cell_type(
                workbook_sha256="abc123",
                address="Sheet1!A1",
                normalized_formula_sha256="def456",
                limits_fingerprint="lim0",
                current_leaf_env={},
            )
            assert result is None
        finally:
            cache.close()

    def test_hit_on_exact_key_match(self, tmp_path: Path) -> None:
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2})))
        leaf_env = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5))
        }
        consumed_keys = ["Sheet1!B1"]
        fp = _compute_leaf_env_subset_fingerprint(consumed_keys, leaf_env)

        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            cache.put_formula_cell_type(
                workbook_sha256="abc123",
                address="Sheet1!A1",
                normalized_formula_sha256="def456",
                limits_fingerprint="lim0",
                leaf_env_subset_fingerprint=fp,
                consumed_leaf_keys=consumed_keys,
                cell_type=ct,
            )
            cache.flush()
            result = cache.get_formula_cell_type(
                workbook_sha256="abc123",
                address="Sheet1!A1",
                normalized_formula_sha256="def456",
                limits_fingerprint="lim0",
                current_leaf_env=leaf_env,
            )
            assert result is not None
            result_ct, result_keys = result
            assert result_ct == ct
            assert result_keys == consumed_keys
        finally:
            cache.close()

    def test_miss_on_workbook_hash_mismatch(self, tmp_path: Path) -> None:
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))
        fp = _compute_leaf_env_subset_fingerprint([], {})
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            cache.put_formula_cell_type(
                workbook_sha256="abc123",
                address="Sheet1!A1",
                normalized_formula_sha256="def456",
                limits_fingerprint="lim0",
                leaf_env_subset_fingerprint=fp,
                consumed_leaf_keys=[],
                cell_type=ct,
            )
            cache.flush()
            result = cache.get_formula_cell_type(
                workbook_sha256="DIFFERENT",
                address="Sheet1!A1",
                normalized_formula_sha256="def456",
                limits_fingerprint="lim0",
                current_leaf_env={},
            )
            assert result is None
        finally:
            cache.close()

    def test_miss_on_formula_hash_mismatch(self, tmp_path: Path) -> None:
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))
        fp = _compute_leaf_env_subset_fingerprint([], {})
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            cache.put_formula_cell_type(
                workbook_sha256="abc",
                address="Sheet1!A1",
                normalized_formula_sha256="formula_v1",
                limits_fingerprint="lim0",
                leaf_env_subset_fingerprint=fp,
                consumed_leaf_keys=[],
                cell_type=ct,
            )
            cache.flush()
            result = cache.get_formula_cell_type(
                workbook_sha256="abc",
                address="Sheet1!A1",
                normalized_formula_sha256="formula_v2",
                limits_fingerprint="lim0",
                current_leaf_env={},
            )
            assert result is None
        finally:
            cache.close()

    def test_miss_on_limits_fingerprint_mismatch(self, tmp_path: Path) -> None:
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))
        fp = _compute_leaf_env_subset_fingerprint([], {})
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            cache.put_formula_cell_type(
                workbook_sha256="abc",
                address="Sheet1!A1",
                normalized_formula_sha256="f1",
                limits_fingerprint="limits_v1",
                leaf_env_subset_fingerprint=fp,
                consumed_leaf_keys=[],
                cell_type=ct,
            )
            cache.flush()
            result = cache.get_formula_cell_type(
                workbook_sha256="abc",
                address="Sheet1!A1",
                normalized_formula_sha256="f1",
                limits_fingerprint="limits_v2",
                current_leaf_env={},
            )
            assert result is None
        finally:
            cache.close()

    def test_hit_on_second_variant_for_same_partial_key(self, tmp_path: Path) -> None:
        """When multiple rows share a partial key but differ in leaf-env
        fingerprint, a lookup with the second variant's env must still hit."""
        ct1 = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))
        ct2 = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({99})))
        env1 = {"Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5))}
        env2 = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=10, max=20))
        }
        keys = ["Sheet1!B1"]
        fp1 = _compute_leaf_env_subset_fingerprint(keys, env1)
        fp2 = _compute_leaf_env_subset_fingerprint(keys, env2)
        assert fp1 != fp2  # sanity

        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            cache.put_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f1",
                limits_fingerprint="lim0",
                leaf_env_subset_fingerprint=fp1,
                consumed_leaf_keys=keys,
                cell_type=ct1,
            )
            cache.put_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f1",
                limits_fingerprint="lim0",
                leaf_env_subset_fingerprint=fp2,
                consumed_leaf_keys=keys,
                cell_type=ct2,
            )
            cache.flush()
            # Look up with env2 — must find ct2
            result = cache.get_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f1",
                limits_fingerprint="lim0",
                current_leaf_env=env2,
            )
            assert result is not None, "Should find the second variant"
            assert result[0] == ct2
        finally:
            cache.close()

    def test_miss_on_leaf_env_fingerprint_mismatch(self, tmp_path: Path) -> None:
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))
        old_env = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5))
        }
        consumed_keys = ["Sheet1!B1"]
        fp = _compute_leaf_env_subset_fingerprint(consumed_keys, old_env)

        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            cache.put_formula_cell_type(
                workbook_sha256="abc",
                address="Sheet1!A1",
                normalized_formula_sha256="f1",
                limits_fingerprint="lim0",
                leaf_env_subset_fingerprint=fp,
                consumed_leaf_keys=consumed_keys,
                cell_type=ct,
            )
            cache.flush()
            # Change the leaf constraint
            new_env = {
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=99, max=99))
            }
            result = cache.get_formula_cell_type(
                workbook_sha256="abc",
                address="Sheet1!A1",
                normalized_formula_sha256="f1",
                limits_fingerprint="lim0",
                current_leaf_env=new_env,
            )
            assert result is None
        finally:
            cache.close()


# ---------------------------------------------------------------------------
# Phase A – Fingerprinting
# ---------------------------------------------------------------------------


class TestFingerprinting:
    def test_leaf_env_fingerprint_stable_across_insertion_order(self) -> None:
        env_a = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5)),
            "Sheet1!C1": CellType(kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"x"}))),
        }
        env_b = {
            "Sheet1!C1": CellType(kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"x"}))),
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5)),
        }
        keys = ["Sheet1!B1", "Sheet1!C1"]
        fp_a = _compute_leaf_env_subset_fingerprint(keys, env_a)
        fp_b = _compute_leaf_env_subset_fingerprint(keys, env_b)
        assert fp_a == fp_b

    def test_unrelated_leaf_env_change_does_not_invalidate(self) -> None:
        keys = ["Sheet1!B1"]
        env_1 = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5)),
            "Sheet1!Z99": CellType(
                kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"old"}))
            ),
        }
        env_2 = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5)),
            "Sheet1!Z99": CellType(
                kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"new"}))
            ),
        }
        assert _compute_leaf_env_subset_fingerprint(
            keys, env_1
        ) == _compute_leaf_env_subset_fingerprint(keys, env_2)

    def test_relevant_leaf_change_changes_fingerprint(self) -> None:
        keys = ["Sheet1!B1"]
        env_1 = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5)),
        }
        env_2 = {
            "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=10)),
        }
        assert _compute_leaf_env_subset_fingerprint(
            keys, env_1
        ) != _compute_leaf_env_subset_fingerprint(keys, env_2)

    def test_limits_fingerprint_changes_with_limits(self) -> None:
        lim1 = DynamicRefLimits(max_branches=1024, max_cells=10000, max_depth=10)
        lim2 = DynamicRefLimits(max_branches=512, max_cells=10000, max_depth=10)
        assert _compute_limits_fingerprint(lim1) != _compute_limits_fingerprint(lim2)

    def test_limits_fingerprint_stable(self) -> None:
        lim = DynamicRefLimits(max_branches=1024, max_cells=10000, max_depth=10)
        assert _compute_limits_fingerprint(lim) == _compute_limits_fingerprint(lim)

    def test_key_hash_deterministic(self) -> None:
        h1 = _compute_key_hash("wb1", "Sheet1!A1", "f1", "lim1", "leaf1")
        h2 = _compute_key_hash("wb1", "Sheet1!A1", "f1", "lim1", "leaf1")
        assert h1 == h2

    def test_key_hash_changes_with_any_field(self) -> None:
        base = _compute_key_hash("wb1", "Sheet1!A1", "f1", "lim1", "leaf1")
        assert base != _compute_key_hash("wb2", "Sheet1!A1", "f1", "lim1", "leaf1")
        assert base != _compute_key_hash("wb1", "Sheet1!A2", "f1", "lim1", "leaf1")
        assert base != _compute_key_hash("wb1", "Sheet1!A1", "f2", "lim1", "leaf1")
        assert base != _compute_key_hash("wb1", "Sheet1!A1", "f1", "lim2", "leaf1")
        assert base != _compute_key_hash("wb1", "Sheet1!A1", "f1", "lim1", "leaf2")


# ---------------------------------------------------------------------------
# Phase B – Incremental durability
# ---------------------------------------------------------------------------


class TestIncrementalDurability:
    def test_write_survives_reopen(self, tmp_path: Path) -> None:
        db_path = tmp_path / "test.sqlite3"
        ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({42})))
        fp = _compute_leaf_env_subset_fingerprint([], {})
        cache = TypeAnalysisCache.open(db_path)
        try:
            cache.put_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f",
                limits_fingerprint="l",
                leaf_env_subset_fingerprint=fp,
                consumed_leaf_keys=[],
                cell_type=ct,
            )
            cache.flush()
        finally:
            cache.close()

        cache2 = TypeAnalysisCache.open(db_path)
        try:
            result = cache2.get_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f",
                limits_fingerprint="l",
                current_leaf_env={},
            )
            assert result is not None
            assert result[0] == ct
        finally:
            cache2.close()

    def test_batched_writes_flush_correctly(self, tmp_path: Path) -> None:
        fp = _compute_leaf_env_subset_fingerprint([], {})
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            for i in range(5):
                cache.put_formula_cell_type(
                    workbook_sha256="wb",
                    address=f"Sheet1!A{i + 1}",
                    normalized_formula_sha256=f"f{i}",
                    limits_fingerprint="l",
                    leaf_env_subset_fingerprint=fp,
                    consumed_leaf_keys=[],
                    cell_type=CellType(
                        kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({i}))
                    ),
                )
            cache.flush()
            for i in range(5):
                result = cache.get_formula_cell_type(
                    workbook_sha256="wb",
                    address=f"Sheet1!A{i + 1}",
                    normalized_formula_sha256=f"f{i}",
                    limits_fingerprint="l",
                    current_leaf_env={},
                )
                assert result is not None
                assert result[0] == CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({i}))
                )
        finally:
            cache.close()

    def test_corrupt_file_degrades_safely(self, tmp_path: Path) -> None:
        db_path = tmp_path / "test.sqlite3"
        db_path.write_text("THIS IS NOT A SQLITE FILE")
        cache = TypeAnalysisCache.open(db_path)
        try:
            result = cache.get_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f",
                limits_fingerprint="l",
                current_leaf_env={},
            )
            assert result is None
            cache.put_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f",
                limits_fingerprint="l",
                leaf_env_subset_fingerprint="e",
                consumed_leaf_keys=[],
                cell_type=CellType(kind=CellKind.ANY),
            )
            cache.flush()  # should not crash
        finally:
            cache.close()

    def test_row_count_cap_evicts_oldest(self, tmp_path: Path) -> None:
        fp = _compute_leaf_env_subset_fingerprint([], {})
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3", max_rows=5)
        try:
            for i in range(10):
                cache.put_formula_cell_type(
                    workbook_sha256="wb",
                    address=f"Sheet1!A{i + 1}",
                    normalized_formula_sha256=f"f{i}",
                    limits_fingerprint="l",
                    leaf_env_subset_fingerprint=fp,
                    consumed_leaf_keys=[],
                    cell_type=CellType(
                        kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({i}))
                    ),
                )
            cache.flush()
            # Oldest entries (A1-A5) should be evicted; newest (A6-A10) should survive
            for i in range(5):
                result = cache.get_formula_cell_type(
                    workbook_sha256="wb",
                    address=f"Sheet1!A{i + 1}",
                    normalized_formula_sha256=f"f{i}",
                    limits_fingerprint="l",
                    current_leaf_env={},
                )
                assert result is None, f"Entry A{i + 1} should have been evicted"
            for i in range(5, 10):
                result = cache.get_formula_cell_type(
                    workbook_sha256="wb",
                    address=f"Sheet1!A{i + 1}",
                    normalized_formula_sha256=f"f{i}",
                    limits_fingerprint="l",
                    current_leaf_env={},
                )
                assert result is not None, f"Entry A{i + 1} should survive eviction"
        finally:
            cache.close()

    def test_diagnostics_counters(self, tmp_path: Path) -> None:
        fp = _compute_leaf_env_subset_fingerprint([], {})
        cache = TypeAnalysisCache.open(tmp_path / "test.sqlite3")
        try:
            ct = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))
            cache.put_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f",
                limits_fingerprint="l",
                leaf_env_subset_fingerprint=fp,
                consumed_leaf_keys=[],
                cell_type=ct,
            )
            cache.flush()
            assert cache.stats.writes == 1

            cache.get_formula_cell_type(
                workbook_sha256="wb",
                address="Sheet1!A1",
                normalized_formula_sha256="f",
                limits_fingerprint="l",
                current_leaf_env={},
            )
            assert cache.stats.hits == 1

            cache.get_formula_cell_type(
                workbook_sha256="NOPE",
                address="Sheet1!A1",
                normalized_formula_sha256="f",
                limits_fingerprint="l",
                current_leaf_env={},
            )
            assert cache.stats.misses == 1
        finally:
            cache.close()
