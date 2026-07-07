"""Tests for export-runtime core delegation in ``emit_runtime``."""

from __future__ import annotations

from typing import Any, cast

import pytest

from excel_grapher.exporter.embed import emit_runtime


def test_emit_runtime_includes_core_math_helpers() -> None:
    code = emit_runtime({"xl_averageif", "xl_sum"}, include_offset_table=False)
    assert "def sum_cells(" in code
    assert "def averageif_cells(" in code
    assert "def xl_averageif(" in code
    assert "sum_cells(*args)" in code
    assert "raise_if_sentinel_float(" in code
    assert "averageif_cells(" in code


def test_emit_runtime_includes_core_abs_helper() -> None:
    code = emit_runtime({"xl_abs"}, include_offset_table=False)
    assert "def abs_number(" in code
    assert "def xl_abs(" in code
    assert "abs_number(*args)" in code
    assert "raise_if_sentinel_float(" in code


def test_emit_runtime_includes_core_text_value_fallback() -> None:
    code = emit_runtime({"xl_value", "xl_numbervalue"}, include_offset_table=False)
    assert "def numbervalue_parse(" in code
    assert "def xl_value(" in code
    assert "numbervalue_parse(text)" in code
    assert "raise_if_sentinel_float(to_number(text))" in code


def test_core_averageif_returns_error_sentinel_without_raising() -> None:
    """Core helpers keep sentinel semantics for evaluator parity."""
    code = emit_runtime({"xl_averageif"}, include_offset_table=False)
    ns: dict[str, Any] = {}
    exec(code, ns)
    core_impl = ns["averageif_cells"]
    result = core_impl([1, 2], ">5", [10, 20, 30])
    assert result == ns["XlError"].VALUE


def test_export_averageif_wrapper_raises_for_same_input() -> None:
    code = emit_runtime({"xl_averageif"}, include_offset_table=False)
    ns: dict[str, Any] = {}
    exec(code, ns)
    wrapper = ns["xl_averageif"]
    xl_error_exception = cast("type[BaseException]", ns["XlErrorException"])
    with pytest.raises(xl_error_exception) as exc_info:
        wrapper([1, 2], ">5", [10, 20, 30])
    assert cast(Any, exc_info.value).code == ns["XlError"].VALUE
