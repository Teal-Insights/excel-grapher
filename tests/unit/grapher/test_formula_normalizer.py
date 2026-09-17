"""Tests for FormulaNormalizer — single-pass named-range substitution and caching."""

from __future__ import annotations

from collections.abc import Callable
from re import Match, Pattern

import pytest

from excel_grapher.core import formula_normalization as formula_normalization_mod
from excel_grapher.grapher import parser as parser_mod
from excel_grapher.grapher.parser import FormulaNormalizer


def _count_re_sub(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    """Install a `re.sub` spy and return the call counter."""
    calls = {"n": 0}
    original = formula_normalization_mod.re.sub

    def counting(
        pattern: str | Pattern[str],
        repl: str | Callable[[Match[str]], str],
        string: str,
        count: int = 0,
        flags: int = 0,
    ) -> str:
        calls["n"] += 1
        return original(pattern, repl, string, count=count, flags=flags)

    monkeypatch.setattr(formula_normalization_mod.re, "sub", counting)
    return calls


class TestFormulaNormalizerBasicNormalization:
    """FormulaNormalizer must produce the regex-normalized formula dialect."""

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

    def test_local_range_qualified(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=SUM(A1:A3)", "Sheet1") == "=SUM(Sheet1!A1:A3)"

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
        assert n.normalize("=SUM(MyTable)", "Sheet1") == "=SUM(Sheet1!A1:B3)"

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

    def test_repeated_calls_skip_the_normalize_pipeline(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Cache hits must not re-enter named-range substitution."""
        named_ranges = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(100)}
        n = FormulaNormalizer(named_ranges=named_ranges)
        formula = "=Name50+Name99"
        compute_ops = {"n": 0}
        original = parser_mod.normalize_excel_formula_with_name_state

        def counting(
            formula: str,
            current_sheet: str,
            *,
            replacements: dict[str, str],
            names_re: object,
        ) -> str:
            compute_ops["n"] += 1
            return original(
                formula,
                current_sheet,
                replacements=replacements,
                names_re=names_re,
            )

        monkeypatch.setattr(parser_mod, "normalize_excel_formula_with_name_state", counting)
        expected = n.normalize(formula, "Sheet1")
        assert expected == "=Sheet1!A51+Sheet1!A100"
        assert compute_ops["n"] == 1
        for _ in range(1000):
            assert n.normalize(formula, "Sheet1") == expected
        assert compute_ops["n"] == 1

    def test_unique_formulas_use_one_name_regex_sub(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Name substitution must not call `re.sub` once per catalog name."""
        calls = _count_re_sub(monkeypatch)

        def re_sub_ops(n_names: int, n_formulas: int) -> int:
            named_ranges = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(n_names)}
            n = FormulaNormalizer(named_ranges=named_ranges)
            calls["n"] = 0
            for i in range(n_formulas):
                assert n.normalize(f"=Name{i}*2", "Sheet1") == f"=Sheet1!A{i + 1}*2"
            return calls["n"]

        assert re_sub_ops(50, 10) == re_sub_ops(200, 10)


class TestFormulaNormalizerOutlierFormula:
    """Regression for a short PV_ResFin formula with a large defined-name catalog."""

    def test_short_formula_with_quoted_sheet_uses_one_name_pass(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Quoted cross-sheet formulas take one name-regex pass (issue #60).

        The LIC-DSF-scale catalog (56 cell names + 39 range names) must not
        walk names one-by-one when the formula mentions none of them.
        """
        formula = "=+'Input 7 - Residual Financing'!$G$14"
        current_sheet = "PV_ResFin-add.int.cost - mkt"
        named_ranges = {f"CellName{i}": ("DataSheet", f"B{i + 1}") for i in range(56)}
        named_range_ranges = {
            f"RangeName{i}": ("DataSheet", f"C{i + 1}", f"D{i + 10}") for i in range(39)
        }
        calls = _count_re_sub(monkeypatch)

        def normalize_with(
            named_ranges: dict[str, tuple[str, str]],
            named_range_ranges: dict[str, tuple[str, str, str]],
        ) -> tuple[str, int]:
            n = FormulaNormalizer(
                named_ranges=named_ranges,
                named_range_ranges=named_range_ranges,
            )
            calls["n"] = 0
            result = n.normalize(formula, current_sheet)
            return result, calls["n"]

        result, large_ops = normalize_with(named_ranges, named_range_ranges)
        tiny_result, tiny_ops = normalize_with({"Unrelated": ("DataSheet", "A1")}, {})
        assert result == tiny_result == "=+'Input 7 - Residual Financing'!G14"
        assert large_ops == tiny_ops
