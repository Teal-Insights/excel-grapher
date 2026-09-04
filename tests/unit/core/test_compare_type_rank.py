"""Scalar comparison uses Excel type-rank (number < text < logical)."""

from __future__ import annotations

from excel_grapher.core.operators import xl_eq, xl_gt, xl_lt
from excel_grapher.core.operators_reference import compare_scalars


def test_compare_scalars_type_rank_matches_excel() -> None:
    assert compare_scalars(">", True, 100) is True
    assert compare_scalars("=", "10", 10) is False
    assert compare_scalars("=", "", 0) is False
    assert compare_scalars("=", "abc", "ABC") is True
    assert compare_scalars("<", "a", 1) is False
    assert compare_scalars("=", True, 1) is False
    assert compare_scalars("=", None, 0) is True


def test_xl_compare_wrappers_follow_type_rank() -> None:
    assert xl_gt(True, 100) is True
    assert xl_eq("10", 10) is False
    assert xl_eq("", 0) is False
    assert xl_lt("a", 1) is False
    assert xl_eq(True, 1) is False
