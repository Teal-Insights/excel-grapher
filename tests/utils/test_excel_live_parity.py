"""Unit tests for live Excel parity helpers (comparison logic; no automation required)."""

from __future__ import annotations

from excel_grapher.core.types import XlError
from tests.utils.excel_live_parity import (
    LiveExcelParityMismatchKind,
    compare_cached_to_evaluator,
)


def test_compare_cached_numeric_match() -> None:
    assert compare_cached_to_evaluator(3.0, 3.0, rtol=1e-5, atol=1e-9) is None


def test_compare_cached_numeric_drift() -> None:
    assert (
        compare_cached_to_evaluator(3.0, 4.0, rtol=1e-5, atol=1e-9)
        == LiveExcelParityMismatchKind.NUMERIC_DRIFT
    )


def test_compare_cached_error_string_matches_xl_error() -> None:
    assert compare_cached_to_evaluator("#NUM!", XlError.NUM, rtol=1e-5, atol=1e-9) is None


def test_compare_cached_error_code_mismatch() -> None:
    assert (
        compare_cached_to_evaluator("#NUM!", XlError.DIV, rtol=1e-5, atol=1e-9)
        == LiveExcelParityMismatchKind.XL_ERROR_CODE_MISMATCH
    )


def test_compare_cached_number_vs_evaluator_error() -> None:
    assert (
        compare_cached_to_evaluator(1.0, XlError.NUM, rtol=1e-5, atol=1e-9)
        == LiveExcelParityMismatchKind.NUMBER_VS_XL_ERROR
    )


def test_compare_cached_error_vs_evaluator_number() -> None:
    assert (
        compare_cached_to_evaluator("#NUM!", 1.0, rtol=1e-5, atol=1e-9)
        == LiveExcelParityMismatchKind.XL_ERROR_VS_NUMBER
    )


def test_compare_cached_string_result() -> None:
    assert compare_cached_to_evaluator("Within 1σ", "Within 1σ", rtol=1e-5, atol=1e-9) is None
