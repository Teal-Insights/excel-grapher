"""Tests for Excel compatibility-prefix handling in grapher utilities."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.builder import _VOLATILE_DYNAMIC_REF_PATTERN


@pytest.mark.parametrize(
    "formula",
    [
        "=NOW()",
        "=_xlfn.NOW()",
        "=_XLUDF.NOW()",
        "=TODAY()",
        "=_xlfn.TODAY()",
        "=RAND()",
        "=_xlfn.RAND()",
        "=RANDBETWEEN(1, 10)",
        "=_xlfn.RANDBETWEEN(1, 10)",
        '=INFO("version")',
        '=_xlfn.INFO("version")',
    ],
)
def test_volatile_pattern_matches_prefixed_and_bare_spellings(formula: str) -> None:
    assert _VOLATILE_DYNAMIC_REF_PATTERN.search(formula.upper()) is not None


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(A1:A10)",
        "=IF(A1>0, B1, C1)",
        "=_xlfn.SUM(A1:A10)",
    ],
)
def test_volatile_pattern_does_not_match_non_volatile_functions(formula: str) -> None:
    assert _VOLATILE_DYNAMIC_REF_PATTERN.search(formula.upper()) is None
