"""End-to-end operator performance budgets and large-range parity.

Loose time ceilings catch catastrophic regressions in CI (default pytest selection).
Micro-benchmark speedups vs ``tests/fixtures/operators_baseline/baseline.json`` are
tracked in the opt-in ``@pytest.mark.slow`` modules:

- ``test_operators_arithmetic_fastpath`` — ``xl_mul`` ~11x
- ``test_operators_compare_fastpath`` — ``xl_gt`` ~5x, ``xl_eq`` ~1.5x
- ``test_operators_concat_fastpath`` — ``xl_concat`` ~2.7x

Perf budget assertions require the optional ``fast`` extra (NumPy). Correctness /
parity cases still run without it.
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.numpy_support import HAS_NUMPY
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.unit.gaps.workbook_helpers import write_large_string_criteria_sumproduct

LARGE_CRITERIA_ROWS = 10_000
LARGE_GRAPH_MAX_RANGE_CELLS = 20_000
SUMPRODUCT_10K_EVAL_TIME_BUDGET_SEC = 5.0
SUMPRODUCT_CHAIN_1K_BASELINE_SPEEDUP_FACTOR = 1.25


@pytest.mark.skipif(not HAS_NUMPY, reason="requires excel-grapher[fast] (NumPy)")
def test_large_string_criteria_sumproduct_evaluator_under_time_budget(
    tmp_path: Path,
) -> None:
    """10K-cell criteria ``SUMPRODUCT`` via ``FormulaEvaluator`` stays within budget."""
    workbook = write_large_string_criteria_sumproduct(
        tmp_path / "criteria_10k.xlsx",
        rows=LARGE_CRITERIA_ROWS,
    )
    graph = create_dependency_graph(
        workbook,
        ["Data!C1"],
        load_values=True,
        use_cached_dynamic_refs=True,
        max_range_cells=LARGE_GRAPH_MAX_RANGE_CELLS,
    )
    started = time.perf_counter()
    with FormulaEvaluator(graph) as evaluator:
        result = evaluator.evaluate("Data!C1")
    elapsed = time.perf_counter() - started
    expected = (LARGE_CRITERIA_ROWS // 2) * 500.0
    assert result == pytest.approx(expected)
    assert elapsed < SUMPRODUCT_10K_EVAL_TIME_BUDGET_SEC, (
        f"10K SUMPRODUCT evaluate took {elapsed:.2f}s, "
        f"expected < {SUMPRODUCT_10K_EVAL_TIME_BUDGET_SEC}s"
    )


def test_large_string_criteria_sumproduct_10k_eval_codegen_parity(tmp_path: Path) -> None:
    """Evaluator and export agree on 10K string-criteria ``SUMPRODUCT``."""
    workbook = write_large_string_criteria_sumproduct(
        tmp_path / "criteria_10k_parity.xlsx",
        rows=LARGE_CRITERIA_ROWS,
    )
    graph = create_dependency_graph(
        workbook,
        ["Data!C1"],
        load_values=True,
        use_cached_dynamic_refs=True,
        max_range_cells=LARGE_GRAPH_MAX_RANGE_CELLS,
    )
    result = assert_codegen_matches_evaluator(graph, ["Data!C1"])
    expected = (LARGE_CRITERIA_ROWS // 2) * 500.0
    assert result.evaluator_results["Data!C1"] == pytest.approx(expected)
    assert result.generated_results["Data!C1"] == pytest.approx(expected)


@pytest.mark.slow
@pytest.mark.skipif(not HAS_NUMPY, reason="requires excel-grapher[fast] (NumPy)")
def test_sumproduct_criteria_chain_1k_beats_baseline() -> None:
    """Direct criteria-chain workload should exceed the Sprint 0 loop baseline."""
    from tests.bench.operators_bench import bench_workload, build_workloads, load_baseline_document
    from tests.unit.core.test_operators_baseline import BASELINE_PATH

    workload = next(w for w in build_workloads() if w.name == "sumproduct_criteria_chain_1k")
    baseline_doc = load_baseline_document(BASELINE_PATH)["workloads"]
    baseline_cps = next(
        entry["cells_per_sec"]
        for entry in baseline_doc
        if entry["name"] == "sumproduct_criteria_chain_1k"
    )
    result = bench_workload(workload, warmup_rounds=2, bench_rounds=5)
    assert result.cells_per_sec >= baseline_cps * SUMPRODUCT_CHAIN_1K_BASELINE_SPEEDUP_FACTOR, (
        f"sumproduct criteria chain 1k: {result.cells_per_sec:.0f} cells/s, "
        f"expected >= {baseline_cps * SUMPRODUCT_CHAIN_1K_BASELINE_SPEEDUP_FACTOR:.0f}"
    )
