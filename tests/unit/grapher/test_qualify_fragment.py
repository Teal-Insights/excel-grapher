"""Tests for `_qualify_fragment` named-range substitution (issue #527)."""

from __future__ import annotations

import re
import time
from collections.abc import Mapping

import pytest

from excel_grapher.grapher import dynamic_refs as dynamic_refs_mod
from excel_grapher.grapher.dynamic_refs import _qualify_fragment


def _qualify(
    expr: str,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
) -> str:
    return _qualify_fragment(expr, named_ranges or {}, named_range_ranges)


class TestQualifyFragmentBehavior:
    """Substitution must stay equivalent to the original per-name `re.sub` loop."""

    def test_empty_and_whitespace_unchanged(self) -> None:
        named = {"Lang": ("Sheet1", "C1")}
        assert _qualify("", named) == ""
        assert _qualify("   ", named) == "   "

    def test_single_cell_name_becomes_sheet_qualified_ref(self) -> None:
        named = {"Lang": ("Sheet1", "C1")}
        assert _qualify("OFFSET(B1,0,Lang)", named) == "OFFSET(B1,0,Sheet1!C1)"

    def test_range_name_dual_qualifies_both_endpoints(self) -> None:
        ranges = {"Country_list": ("lookup", "C4", "C6")}
        assert _qualify("INDEX(Country_list,1,1)", named_range_ranges=ranges) == (
            "INDEX(lookup!C4:lookup!C6,1,1)"
        )

    def test_quoted_sheet_name_in_replacement(self) -> None:
        named = {"Input": ("My Sheet", "A1")}
        assert _qualify("Input*2", named) == "'My Sheet'!A1*2"

    def test_name_used_as_sheet_qualifier_is_not_replaced(self) -> None:
        named = {"Foo": ("Other", "C5")}
        assert _qualify("Foo!A1+Foo", named) == "Foo!A1+Other!C5"

    def test_longer_name_wins_over_dotted_prefix(self) -> None:
        """`.` is a non-word character, so longest-first ordering is load-bearing."""
        named = {
            "Sales": ("Sheet1", "A1"),
            "Sales.Total": ("Sheet1", "B2"),
        }
        assert _qualify("Sales.Total+Sales", named) == "Sheet1!B2+Sheet1!A1"

    def test_cell_map_wins_when_name_is_in_both_maps(self) -> None:
        named = {"X": ("Sheet1", "A1")}
        ranges = {"X": ("Sheet1", "B1", "B3")}
        assert _qualify("X", named, ranges) == "Sheet1!A1"

    def test_absent_names_do_not_change_fragment(self) -> None:
        named = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(50)}
        assert _qualify("1+2", named) == "1+2"


class TestQualifyFragmentRegexCache:
    """Issue #527: do not recompile a string pattern per defined name per call."""

    def test_does_not_pass_string_patterns_into_re_sub_per_name(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        string_pattern_subs: list[str] = []
        real_sub = dynamic_refs_mod.re.sub

        def counting_sub(
            pattern: str | re.Pattern[str],
            repl: str,
            string: str,
            count: int = 0,
            flags: int = 0,
        ) -> str:
            if isinstance(pattern, str):
                string_pattern_subs.append(pattern)
            return real_sub(pattern, repl, string, count=count, flags=flags)

        monkeypatch.setattr(dynamic_refs_mod.re, "sub", counting_sub)

        named = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(300)}
        result = _qualify("Name42+1", named)
        assert result == "Sheet1!A43+1"
        # One present name at most; compiling a string per catalog entry is the bug.
        assert len(string_pattern_subs) <= 1, (
            f"re.sub received {len(string_pattern_subs)} string patterns; "
            "patterns must be precompiled and absent names skipped"
        )

    def test_token_patterns_are_lru_cached_across_calls(self) -> None:
        pattern_fn = dynamic_refs_mod._defined_name_token_pattern
        pattern_fn.cache_clear()
        named = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(20)}
        for _ in range(5):
            _qualify("Name3+Name7", named)
        info = pattern_fn.cache_info()
        assert info.misses <= 2
        assert info.hits >= 8
        assert info.maxsize is not None

    def test_large_catalog_repeat_calls_stay_fast(self) -> None:
        named = {f"Name{i}": ("Sheet1", f"A{i + 1}") for i in range(600)}
        expr = "Name5+Name9"
        # Warm compiled patterns / replacement pairs.
        _qualify(expr, named)
        start = time.perf_counter()
        for _ in range(300):
            assert _qualify(expr, named) == "Sheet1!A6+Sheet1!A10"
        elapsed = time.perf_counter() - start
        assert elapsed < 0.5, f"qualify_fragment too slow with 600 names: {elapsed:.3f}s"
