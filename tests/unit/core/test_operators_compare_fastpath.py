"""Vectorized comparison fast paths for binary operators."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.core.operators import xl_eq
from excel_grapher.core.types import XlError
from tests.bench.operators_bench import (
    bench_workload,
    build_workloads,
    category_column,
    load_baseline_document,
    numeric_column,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.unit.core.operators_test_helpers import (
    COMPARE_OPS,
    as_ndarray,
    assert_compare_matches_reference,
)
from tests.unit.core.test_operators_baseline import BASELINE_PATH
from tests.unit.gaps.workbook_helpers import write_large_string_criteria_sumproduct

LARGE_SHAPE = (10_000, 1)
MEDIUM_SHAPE = (1_000, 1)
EQ_STRING_10K_BASELINE_SPEEDUP_FACTOR = 1.25
GT_NUMERIC_1K_BASELINE_SPEEDUP_FACTOR = 5.0


@pytest.mark.parametrize("op", COMPARE_OPS)
def test_compare_fastpath_matches_reference_on_large_string_equality(op: str) -> None:
    categories = category_column(LARGE_SHAPE, seed=31)
    assert_compare_matches_reference(op, categories, "Software")


@pytest.mark.parametrize("op", [">", ">=", "<", "<="])
def test_compare_fastpath_matches_reference_on_large_numeric_threshold(op: str) -> None:
    numbers = numeric_column(MEDIUM_SHAPE, seed=41)
    assert_compare_matches_reference(op, numbers, 200)


def test_compare_fastpath_string_equality_is_case_insensitive_at_scale() -> None:
    labels = np.array([["software", "SOFTWARE", "Software"]], dtype=object)
    result = as_ndarray(xl_eq(labels, "software"))
    assert result.tolist() == [[True, True, True]]


def test_compare_fastpath_matches_reference_with_numeric_strings() -> None:
    left = np.array([["10", " 2.5 "], ["0", ""]], dtype=object)
    right = np.array([[10.0, 2.5], [0.0, 0.0]], dtype=object)
    assert_compare_matches_reference("=", left, right)
    assert_compare_matches_reference("<=", left, right)


def test_compare_fastpath_falls_back_on_mixed_type_cells() -> None:
    left = np.array([["TRUE", 2.0]], dtype=object)
    right = np.array([[True, "2"]], dtype=object)
    assert_compare_matches_reference("=", left, right)


def test_compare_fastpath_falls_back_on_non_numeric_string_with_number() -> None:
    left = np.array([["abc", 2.0]], dtype=object)
    right = np.array([[0.0, 2.0]], dtype=object)
    assert_compare_matches_reference("=", left, right)


def test_compare_fastpath_fail_fast_on_first_error_at_index_zero() -> None:
    left = np.array([[XlError.NA, 2.0], [3.0, 4.0]], dtype=object)
    right = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    assert xl_eq(left, right) == XlError.NA


def test_compare_fastpath_fail_fast_on_first_error_late_in_c_order() -> None:
    left = np.array([[1.0, 2.0], [XlError.DIV, 4.0]], dtype=object)
    right = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    assert xl_eq(left, right) == XlError.DIV


def test_compare_fastpath_left_error_wins_over_right_error_per_cell() -> None:
    left = np.array([[XlError.REF, XlError.NA]], dtype=object)
    right = np.array([[XlError.NA, XlError.REF]], dtype=object)
    assert xl_eq(left, right) == XlError.REF


def test_compare_scalar_path_unchanged() -> None:
    assert xl_eq(1, 1) is True
    assert xl_eq("AbC", "aBc") is True
    assert xl_eq(XlError.NA, 0) == XlError.NA


def test_large_string_criteria_sumproduct_eval_codegen_parity(tmp_path: Path) -> None:
    """Evaluator and export agree on large ``SUMPRODUCT`` with string criteria."""
    workbook = write_large_string_criteria_sumproduct(
        tmp_path / "large_string_sumproduct.xlsx",
        rows=2_000,
    )
    graph = create_dependency_graph(
        workbook,
        ["Data!C1"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    result = assert_codegen_matches_evaluator(graph, ["Data!C1"])
    assert result.evaluator_results["Data!C1"] == pytest.approx(500_000.0)
    assert result.generated_results["Data!C1"] == pytest.approx(500_000.0)


@pytest.mark.slow
def test_xl_eq_string_10k_beats_baseline() -> None:
    workload = next(w for w in build_workloads() if w.name == "xl_eq_string_10k")
    baseline_doc = load_baseline_document(BASELINE_PATH)["workloads"]
    baseline_cps = next(
        entry["cells_per_sec"] for entry in baseline_doc if entry["name"] == "xl_eq_string_10k"
    )
    result = bench_workload(workload, warmup_rounds=2, bench_rounds=5)
    assert result.cells_per_sec >= baseline_cps * EQ_STRING_10K_BASELINE_SPEEDUP_FACTOR, (
        f"xl_eq string 10k: {result.cells_per_sec:.0f} cells/s, "
        f"expected >= {baseline_cps * EQ_STRING_10K_BASELINE_SPEEDUP_FACTOR:.0f}"
    )


@pytest.mark.slow
def test_xl_gt_numeric_1k_beats_baseline() -> None:
    workload = next(w for w in build_workloads() if w.name == "xl_gt_numeric_1k")
    baseline_doc = load_baseline_document(BASELINE_PATH)["workloads"]
    baseline_cps = next(
        entry["cells_per_sec"] for entry in baseline_doc if entry["name"] == "xl_gt_numeric_1k"
    )
    result = bench_workload(workload, warmup_rounds=2, bench_rounds=5)
    assert result.cells_per_sec >= baseline_cps * GT_NUMERIC_1K_BASELINE_SPEEDUP_FACTOR, (
        f"xl_gt numeric 1k: {result.cells_per_sec:.0f} cells/s, "
        f"expected >= {baseline_cps * GT_NUMERIC_1K_BASELINE_SPEEDUP_FACTOR:.0f}"
    )
