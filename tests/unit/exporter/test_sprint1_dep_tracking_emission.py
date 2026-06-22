"""Sprint 1: dual cache scaffold extraction for slim vs full ``emit_runtime`` emission."""

from __future__ import annotations

from collections.abc import Callable
from typing import cast

from excel_grapher.core import CellValue
from excel_grapher.exporter.embed import emit_runtime, runtime_cache_seed_symbols
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict, xl_cell
from tests.integration.utils.parity_harness import (
    SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET,
    assert_dep_tracking_absent,
    assert_dep_tracking_present,
    count_cache_eval_scaffold_lines,
    count_dep_tracking_lines,
)


def _minimal_cache_runtime(*, include_dep_tracking: bool) -> str:
    return emit_runtime(
        runtime_cache_seed_symbols(include_dep_tracking=include_dep_tracking),
        include_offset_table=False,
        include_dep_tracking=include_dep_tracking,
    )


def test_runtime_cache_seed_symbols_differs_by_dep_tracking_mode() -> None:
    full = runtime_cache_seed_symbols(include_dep_tracking=True)
    slim = runtime_cache_seed_symbols(include_dep_tracking=False)
    assert full == slim | {"xl_circular_reference"}


def test_emit_runtime_slim_omits_dep_tracking_scaffold() -> None:
    code = _minimal_cache_runtime(include_dep_tracking=False)
    assert_dep_tracking_absent(code)
    assert count_dep_tracking_lines(code) == 0
    assert count_cache_eval_scaffold_lines(code) <= SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET


def test_emit_runtime_full_includes_dep_tracking_scaffold() -> None:
    code = _minimal_cache_runtime(include_dep_tracking=True)
    assert_dep_tracking_present(code)
    assert count_dep_tracking_lines(code) >= 40


def test_slim_emitted_runtime_evaluates_minimal_workbook() -> None:
    code = _minimal_cache_runtime(include_dep_tracking=False)

    namespace: dict[str, object] = {}
    exec(code, namespace)

    eval_context = cast(type[EvalContext], namespace["EvalContext"])
    coerce = cast(
        Callable[[dict[str, object]], dict[str, CellValue]], namespace["coerce_inputs_dict"]
    )
    cell = cast(Callable[[EvalContext, str], CellValue], namespace["xl_cell"])

    def resolver(address: str) -> Callable[[EvalContext], CellValue] | None:
        if address == "S!B1":
            return lambda ctx: cast(float, cell(ctx, "S!A1")) + 1.0
        return None

    ctx = eval_context(inputs=coerce({"S!A1": 1.0}), resolver=resolver)
    assert cell(ctx, "S!B1") == 2.0


def test_slim_and_full_emitted_runtimes_agree_on_minimal_evaluation() -> None:
    slim_ns: dict[str, object] = {}
    full_ns: dict[str, object] = {}
    exec(_minimal_cache_runtime(include_dep_tracking=False), slim_ns)
    exec(_minimal_cache_runtime(include_dep_tracking=True), full_ns)

    def run(namespace: dict[str, object]) -> float:
        eval_context = cast(type[EvalContext], namespace["EvalContext"])
        coerce = cast(
            Callable[[dict[str, object]], dict[str, CellValue]], namespace["coerce_inputs_dict"]
        )
        cell = cast(Callable[[EvalContext, str], CellValue], namespace["xl_cell"])

        def resolver(address: str) -> Callable[[EvalContext], CellValue] | None:
            if address == "S!B1":
                return lambda ctx: cast(float, cell(ctx, "S!A1")) + 1.0
            return None

        ctx = eval_context(inputs=coerce({"S!A1": 3.0}), resolver=resolver)
        result = cell(ctx, "S!B1")
        assert isinstance(result, float)
        return result

    assert run(slim_ns) == run(full_ns) == 4.0


def test_library_eval_context_retains_dep_tracking_after_split() -> None:
    ctx = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _address: None)
    ctx.set_inputs({"S!A1": 1.0})
    assert ctx.inputs["S!A1"] == 1.0


def test_library_xl_cell_still_records_dependencies() -> None:
    child = "S!B1"

    def parent_fn(ctx: EvalContext) -> CellValue:
        xl_cell(ctx, child)
        return 1

    ctx = EvalContext(
        inputs=coerce_inputs_dict({child: 10}),
        resolver=lambda address: parent_fn if address == "S!A1" else None,
    )
    xl_cell(ctx, "S!A1")
    assert child in ctx.deps["S!A1"]
