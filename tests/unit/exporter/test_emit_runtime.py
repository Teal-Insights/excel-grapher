"""Tests for ``emit_runtime`` symbol selection."""

from __future__ import annotations

from excel_grapher.exporter.embed import emit_runtime


def test_emit_runtime_includes_xl_abs_definition() -> None:
    code = emit_runtime({"xl_abs"}, include_offset_table=False)
    assert "def xl_abs" in code
    assert "xl_abs" in code


def test_emit_runtime_includes_smoke_blocker_symbols() -> None:
    code = emit_runtime(
        {
            "xl_averageif",
            "xl_iserror",
            "xl_isna",
            "xl_lower",
            "xl_today",
            "xl_value",
        },
        include_offset_table=False,
    )
    for symbol in (
        "def xl_averageif",
        "def xl_iserror",
        "def xl_isna",
        "def xl_lower",
        "def xl_today",
        "def xl_value",
    ):
        assert symbol in code
