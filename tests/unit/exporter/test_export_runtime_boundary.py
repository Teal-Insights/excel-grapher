"""Tests for export-runtime sentinel shadowing in ``emit_runtime``."""

from __future__ import annotations

from excel_grapher.exporter.embed import emit_runtime
from excel_grapher.exporter.export_runtime.offset import xl_offset_ref


def test_emit_runtime_shadows_sentinel_math_helpers() -> None:
    code = emit_runtime({"xl_averageif", "xl_sum"}, include_offset_table=False)
    assert "def _sentinel_xl_averageif" in code
    assert "def _sentinel_xl_sum" in code
    assert "def xl_averageif" in code
    assert "raise_if_sentinel_float(_sentinel_xl_averageif" in code
    assert "raise_if_sentinel_float(_sentinel_xl_sum" in code


def test_emit_runtime_shadows_sentinel_text_value_fallback() -> None:
    code = emit_runtime({"xl_value", "xl_numbervalue"}, include_offset_table=False)
    assert "def _sentinel_xl_numbervalue" in code
    assert "def xl_value" in code
    assert "_sentinel_xl_numbervalue(text)" in code
    assert "raise_if_sentinel_float(to_number(text))" in code


def test_export_offset_ref_raises_reference_errors() -> None:
    import pytest

    from excel_grapher.core.types import XlErrorException
    from excel_grapher.evaluator.types import XlError

    with pytest.raises(XlErrorException) as exc_info:
        xl_offset_ref(("Sheet1", 1, 1, 3, 1), -5, 0)
    assert exc_info.value.code == XlError.REF
