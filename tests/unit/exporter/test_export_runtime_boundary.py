"""Tests for export-runtime boundary helpers and codegen wrapping."""

from __future__ import annotations

from excel_grapher.exporter.embed import emit_runtime
from excel_grapher.exporter.export_runtime.offset import xl_offset_ref


def test_emit_runtime_includes_raise_if_sentinel() -> None:
    code = emit_runtime({"raise_if_sentinel", "xl_sum"}, include_offset_table=False)
    assert "def raise_if_sentinel" in code
    assert "def xl_sum" in code


def test_export_offset_ref_raises_reference_errors() -> None:
    import pytest

    from excel_grapher.core.types import XlErrorException
    from excel_grapher.evaluator.types import XlError

    with pytest.raises(XlErrorException) as exc_info:
        xl_offset_ref(("Sheet1", 1, 1, 3, 1), -5, 0)
    assert exc_info.value.code == XlError.REF
