"""Relational text compare keeps trailing spaces significant (Excel parity).

GitHub #434 claimed `compare_scalars` should strip trailing spaces so
`"High" = "High "` is True. Live Excel returns False for that expression
(and for cell-vs-cell equality); exact `MATCH(..., 0)` likewise treats
trailing spaces as significant. These tests lock the Excel-aligned behavior
so that incorrect "fix" does not land later.
"""

from __future__ import annotations

import numpy as np
import pytest

from excel_grapher.core.operators import xl_eq, xl_ne
from excel_grapher.core.operators_reference import compare_scalars
from excel_grapher.core.types import XlError
from excel_grapher.runtime.lookup import xl_match
from tests.unit.core.operators_test_helpers import as_ndarray


@pytest.mark.parametrize(
    ("op", "left", "right", "expected"),
    [
        ("=", "High", "High ", False),
        ("<>", "High", "High ", True),
        ("=", " High", "High", False),
        ("=", "High", "High  ", False),
        ("=", "high", "HIGH", True),
        ("=", "high", "HIGH ", False),
        ("=", "High ", "High ", True),
        ("<", "ABC ", "ABB", False),
        (">", "ABC ", "ABB", True),
    ],
)
def test_compare_scalars_trailing_spaces_match_excel(
    op: str, left: str, right: str, expected: bool
) -> None:
    assert compare_scalars(op, left, right) is expected


def test_xl_eq_and_xl_ne_preserve_trailing_spaces() -> None:
    assert xl_eq("High", "High ") is False
    assert xl_ne("High", "High ") is True
    assert xl_eq("High ", "High ") is True
    assert xl_eq("high", "HIGH") is True


def test_array_compare_preserves_trailing_spaces() -> None:
    left = np.array([["High", "High "], [" High", "high"]], dtype=object)
    right = np.array([["High ", "High "], ["High", "HIGH"]], dtype=object)
    assert as_ndarray(xl_eq(left, right)).tolist() == [[False, True], [False, True]]


def test_match_exact_treats_trailing_spaces_as_significant() -> None:
    haystack = np.array([["High "]], dtype=object)
    assert xl_match("High", haystack, 0) == XlError.NA
    assert xl_match("High ", haystack, 0) == 1.0
    assert xl_match("High", np.array([["High"]], dtype=object), 0) == 1.0
