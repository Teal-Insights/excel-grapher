"""Unit tests for live Excel parity helpers (comparison logic; no automation required)."""

from __future__ import annotations

from tests.utils.excel_live_parity import (
    LiveExcelParityMismatchKind,
    compare_cached_to_evaluator,
)


def test_compare_cached_numeric_drift() -> None:
    assert (
        compare_cached_to_evaluator(3.0, 4.0, rtol=1e-5, atol=1e-9)
        == LiveExcelParityMismatchKind.NUMERIC_DRIFT
    )


def test_compare_cached_string_result() -> None:
    assert compare_cached_to_evaluator("Within 1σ", "Within 1σ", rtol=1e-5, atol=1e-9) is None
