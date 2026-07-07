"""Tests for export-runtime sentinel shadowing in ``emit_runtime``."""

from __future__ import annotations

import ast
from pathlib import Path
from typing import Any, cast

import pytest

from excel_grapher.exporter.embed import _format_emitted_symbol_source, emit_runtime

_RUNTIME_MATH = Path(__file__).resolve().parents[3] / "excel_grapher" / "runtime" / "math.py"


def test_format_emitted_symbol_source_renames_shadowed_def() -> None:
    """Sentinel emission renames only the top-level ``def`` for shadowed symbols.

    Invariants exercised here (see ``embed._format_emitted_symbol_source``):

    * The shadowed symbol uses the ``_sentinel_{original}`` prefix.
    * The shared implementation is a plain top-level ``def {original}(...`` with
      no decorators, so a single ``def {original}(`` replacement is sufficient.
    * Internal references to other ``xl_*`` names in the extracted body are left
      untouched by the rename pass.
    """
    module_src = _RUNTIME_MATH.read_text(encoding="utf-8")
    module_ast = ast.parse(module_src, filename=str(_RUNTIME_MATH))
    sum_node = next(
        node
        for node in module_ast.body
        if isinstance(node, ast.FunctionDef) and node.name == "xl_sum"
    )

    segment = _format_emitted_symbol_source("_sentinel_xl_sum", module_src, sum_node)

    assert segment.startswith("def _sentinel_xl_sum(")
    assert "def xl_sum(" not in segment
    assert "numeric_values" in segment


def test_emit_runtime_shadows_sentinel_math_helpers() -> None:
    code = emit_runtime({"xl_averageif", "xl_sum"}, include_offset_table=False)
    assert "def _sentinel_xl_averageif" in code
    assert "def _sentinel_xl_sum" in code
    assert "def xl_averageif" in code
    assert "_sentinel_xl_averageif" in code
    assert "raise_if_sentinel_float(" in code
    assert "_sentinel_xl_sum(*args)" in code


def test_emit_runtime_shadows_core_abs_helper() -> None:
    code = emit_runtime({"xl_abs"}, include_offset_table=False)
    assert "def _sentinel_xl_abs" in code
    assert "def xl_abs" in code
    assert "_sentinel_xl_abs(*args)" in code
    assert "raise_if_sentinel_float(" in code


def test_emit_runtime_shadows_sentinel_text_value_fallback() -> None:
    code = emit_runtime({"xl_value", "xl_numbervalue"}, include_offset_table=False)
    assert "def _sentinel_xl_numbervalue" in code
    assert "def xl_value" in code
    assert "_sentinel_xl_numbervalue(text)" in code
    assert "raise_if_sentinel_float(to_number(text))" in code


def test_sentinel_averageif_returns_error_sentinel_without_raising() -> None:
    """Internal ``_sentinel_*`` helpers keep sentinel semantics for delegation."""
    code = emit_runtime({"xl_averageif"}, include_offset_table=False)
    ns: dict[str, Any] = {}
    exec(code, ns)
    sentinel = ns["_sentinel_xl_averageif"]
    result = sentinel([1, 2], ">5", [10, 20, 30])
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
