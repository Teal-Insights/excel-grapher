"""Tests for FormulaNormalizer — single-pass named-range substitution and caching."""

from __future__ import annotations

import time

from excel_grapher.grapher.parser import FormulaNormalizer, normalize_formula


class TestFormulaNormalizerBasicNormalization:
    """FormulaNormalizer must produce identical output to normalize_formula."""

    def test_same_sheet_ref_qualified(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=A1+A2", "Sheet1") == "=Sheet1!A1+Sheet1!A2"

    def test_strips_absolute_markers(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=$A$1+$A1+A$1", "Sheet1") == "=Sheet1!A1+Sheet1!A1+Sheet1!A1"

    def test_cross_sheet_ref_preserved(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("='Other Sheet'!$B$5+1", "Sheet1") == "='Other Sheet'!B5+1"

    def test_quoted_sheet_redundant_quotes_stripped_to_match_graph_keys(self) -> None:
        """Excel may quote sheet names unnecessarily; keys use format_cell_key rules."""
        n = FormulaNormalizer()
        f = "='C3_commodity_prices_pub'!$E$13"
        expected = "=C3_commodity_prices_pub!E13"
        assert n.normalize(f, "Sheet1") == expected
        assert normalize_formula(f, "Sheet1") == expected

    def test_local_range_qualified(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=SUM(A1:A3)", "Sheet1") == "=SUM(Sheet1!A1:Sheet1!A3)"

    def test_non_formula_returned_unchanged(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("hello", "Sheet1") == "hello"
        assert n.normalize("", "Sheet1") == ""


class TestFormulaNormalizerNamedRanges:
    """Named-range substitution via single-pass alternation regex."""

    def test_single_cell_named_range_resolved(self) -> None:
        named_ranges = {"MyInput": ("Sheet1", "A1")}
        n = FormulaNormalizer(named_ranges=named_ranges)
        assert n.normalize("=MyInput*2", "Sheet1") == "=Sheet1!A1*2"

    def test_range_named_range_resolved(self) -> None:
        named_range_ranges = {"MyTable": ("Sheet1", "A1", "B3")}
        n = FormulaNormalizer(named_range_ranges=named_range_ranges)
        assert n.normalize("=SUM(MyTable)", "Sheet1") == "=SUM(Sheet1!A1:Sheet1!B3)"

    def test_name_inside_sheet_ref_not_replaced(self) -> None:
        """Names appearing as sheet qualifiers (Foo!A1) must not be replaced."""
        named_ranges = {"Foo": ("Other", "C5")}
        n = FormulaNormalizer(named_ranges=named_ranges)
        # Foo!A1 — Foo is the sheet, not the name
        result = n.normalize("=Foo!A1", "Sheet1")
        assert "Other!C5" not in result
        assert "Foo!A1" in result

    def test_longer_name_wins_over_prefix(self) -> None:
        """If names share a prefix, the longer one must match first."""
        named_ranges = {
            "Rate": ("Sheet1", "A1"),
            "RateAdj": ("Sheet1", "B2"),
        }
        n = FormulaNormalizer(named_ranges=named_ranges)
        result = n.normalize("=RateAdj+Rate", "Sheet1")
        assert result == "=Sheet1!B2+Sheet1!A1"

    def test_many_names_only_present_substituted(self) -> None:
        """Names not present in the formula must not corrupt the result."""
        named_ranges = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(50)}
        named_ranges["Target"] = ("Sheet1", "Z1")
        n = FormulaNormalizer(named_ranges=named_ranges)
        result = n.normalize("=Target*2", "Sheet1")
        assert result == "=Sheet1!Z1*2"


class TestFormulaNormalizerCaching:
    """normalize() must return cached results on repeated calls."""

    def test_same_result_returned_on_repeat(self) -> None:
        n = FormulaNormalizer()
        r1 = n.normalize("=A1+A2", "Sheet1")
        r2 = n.normalize("=A1+A2", "Sheet1")
        assert r1 == r2

    def test_cache_is_per_sheet(self) -> None:
        n = FormulaNormalizer()
        r1 = n.normalize("=A1", "Sheet1")
        r2 = n.normalize("=A1", "Sheet2")
        assert r1 == "=Sheet1!A1"
        assert r2 == "=Sheet2!A1"

    def test_large_name_set_fast_on_repeat(self) -> None:
        """With 100 named ranges and 1000 repeated calls, caching keeps it fast."""
        named_ranges = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(100)}
        n = FormulaNormalizer(named_ranges=named_ranges)
        formula = "=Name50+Name99"
        # Warm up
        n.normalize(formula, "Sheet1")
        start = time.perf_counter()
        for _ in range(1000):
            n.normalize(formula, "Sheet1")
        elapsed = time.perf_counter() - start
        # 1000 cached lookups should complete in well under 100ms
        assert elapsed < 0.1, f"Cached calls too slow: {elapsed:.3f}s"

    def test_large_name_set_single_pass_fast(self) -> None:
        """With 100 named ranges, a single normalize call must complete quickly."""
        named_ranges = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(100)}
        n = FormulaNormalizer(named_ranges=named_ranges)
        start = time.perf_counter()
        for _ in range(200):
            # Different formulas so cache doesn't help
            n.normalize(f"=Name{_ % 100}*2", f"Sheet{_ % 5}")
        elapsed = time.perf_counter() - start
        # 200 unique calls with 100 names must complete in under 2s
        assert elapsed < 2.0, f"Single-pass normalization too slow: {elapsed:.3f}s"


class TestFormulaNormalizerOutlierFormula:
    """Regression test for the ~380ms outlier on a short formula (issue #60)."""

    def test_short_formula_with_quoted_sheet_fast(self) -> None:
        """
        Formula from PV_ResFin sheet with quoted cross-ref must normalize quickly
        even with a large set of named ranges present.
        """
        # Simulate 56 cell names + 39 range names (95 total, like the LIC-DSF workbook)
        named_ranges = {f"CellName{i}": ("DataSheet", f"B{i + 1}") for i in range(56)}
        named_range_ranges = {
            f"RangeName{i}": ("DataSheet", f"C{i + 1}", f"D{i + 10}") for i in range(39)
        }
        n = FormulaNormalizer(
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
        formula = "=+'Input 7 - Residual Financing'!$G$14"
        current_sheet = "PV_ResFin-add.int.cost - mkt"

        start = time.perf_counter()
        for _ in range(100):
            n._cache.clear()  # bypass cache to measure raw normalization cost
            n.normalize(formula, current_sheet)
        elapsed = time.perf_counter() - start
        # 100 calls without cache must complete in under 1s (10ms/call budget)
        assert elapsed < 1.0, f"Outlier formula still too slow: {elapsed:.3f}s for 100 calls"
