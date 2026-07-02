"""Tests for ``emit_runtime`` symbol selection."""

from __future__ import annotations

from typing import Any, cast

from excel_grapher.exporter.embed import emit_runtime, runtime_cache_seed_symbols
from tests.integration.utils.parity_harness import (
    assert_dep_tracking_absent,
    assert_dep_tracking_present,
)


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


def test_emit_runtime_dep_tracking_flag_selects_scaffold() -> None:
    slim = emit_runtime(
        runtime_cache_seed_symbols(include_dep_tracking=False),
        include_offset_table=False,
        include_dep_tracking=False,
    )
    full = emit_runtime(
        runtime_cache_seed_symbols(include_dep_tracking=True),
        include_offset_table=False,
        include_dep_tracking=True,
    )
    assert_dep_tracking_absent(slim)
    assert_dep_tracking_present(full)


def test_emit_runtime_operator_symbols_resolve_to_export_runtime() -> None:
    """Exported operators consume lazy ranges; numpy fastpaths stay out of exports."""
    symbols = {
        "xl_compare",
        "xl_number",
        "xl_sumproduct",
        *runtime_cache_seed_symbols(include_dep_tracking=False),
    }
    for include_operators_fastpath in (False, True):
        code = emit_runtime(
            symbols,
            include_offset_table=False,
            include_dep_tracking=False,
            include_operators_fastpath=include_operators_fastpath,
        )
        assert "batch_coerce_to_float64" not in code
        assert "import numpy" not in code
        assert "class Range" in code


def test_emit_runtime_includes_export_runtime_primitives() -> None:
    code = emit_runtime({"Range"}, include_offset_table=False)
    assert "class XlErrorException" in code
    assert "class Range" in code

    ns: dict[str, Any] = {}
    exec(code, ns)
    range_type = ns["Range"]
    xl_error = ns["XlError"]
    xl_error_exception = cast(type[BaseException], ns["XlErrorException"])

    calls: list[str] = []

    def resolve(address: str) -> Any:
        calls.append(address)
        return xl_error.DIV if address == "S!B1" else 1

    rng = range_type("S", 1, 1, 1, 2, resolve)
    try:
        list(rng)
    except xl_error_exception as exc:
        assert cast(Any, exc).code == xl_error.DIV
    else:
        raise AssertionError("Expected exported Range iteration to raise")

    assert calls == ["S!A1", "S!B1"]
