"""Vectorized concat fast path for binary operators."""

from __future__ import annotations

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core.operators import xl_concat
from excel_grapher.core.types import XlError
from tests.bench.operators_bench import (
    bench_workload,
    build_workloads,
    digit_suffix_column,
    load_baseline_document,
    string_prefix_column,
)
from tests.unit.core.operators_test_helpers import array_tolist, assert_concat_matches_reference
from tests.unit.core.test_operators_baseline import BASELINE_PATH

LARGE_SHAPE = (1_000, 1)
CONCAT_1K_BASELINE_SPEEDUP_FACTOR = 1.25


def test_concat_fastpath_matches_reference_on_benchmark_columns() -> None:
    left = string_prefix_column(LARGE_SHAPE, seed=81)
    right = digit_suffix_column(LARGE_SHAPE, seed=82)
    assert_concat_matches_reference(left, right)


def test_concat_fastpath_matches_reference_on_string_columns() -> None:
    left = np.array([["a", "b"], ["c", "d"]], dtype=object)
    right = np.array([["1", "2"], ["3", "4"]], dtype=object)
    assert_concat_matches_reference(left, right)


def test_concat_fastpath_matches_reference_with_scalar_broadcast() -> None:
    left = np.array([["x", "y"] * 50], dtype=object).reshape(100, 1)
    assert_concat_matches_reference(left, "!")


def test_concat_fastpath_matches_reference_with_numeric_columns() -> None:
    left = np.array([[1.0, 2.25], [3.0, 4.5]], dtype=object)
    right = np.array([[10.0, 20.0], [30.0, 40.0]], dtype=object)
    assert_concat_matches_reference(left, right)


def test_concat_fastpath_matches_reference_with_bools_and_none() -> None:
    left = np.array([[True, None], [False, "x"]], dtype=object)
    right = np.array([[1.0, 2.0], ["!", ""]], dtype=object)
    assert_concat_matches_reference(left, right)


def test_concat_fastpath_falls_back_on_mixed_type_column() -> None:
    left = np.array([["a", 1.0], [True, None]], dtype=object)
    right = np.array([["z", 2], ["!", ""]], dtype=object)
    assert_concat_matches_reference(left, right)


def test_concat_fastpath_embedded_error_becomes_string() -> None:
    left = np.array([[XlError.NA, "a"]], dtype=object)
    right = np.array([["!", "b"]], dtype=object)
    assert array_tolist(xl_concat(left, right)) == [["#N/A!", "ab"]]


def test_concat_scalar_path_unchanged() -> None:
    assert xl_concat("a", "b") == "ab"
    assert xl_concat(2.0, 3.0) == "23"


@pytest.mark.slow
def test_xl_concat_string_1k_beats_baseline() -> None:
    workload = next(w for w in build_workloads() if w.name == "xl_concat_string_1k")
    baseline_doc = load_baseline_document(BASELINE_PATH)["workloads"]
    baseline_cps = next(
        entry["cells_per_sec"] for entry in baseline_doc if entry["name"] == "xl_concat_string_1k"
    )
    result = bench_workload(workload, warmup_rounds=2, bench_rounds=5)
    assert result.cells_per_sec >= baseline_cps * CONCAT_1K_BASELINE_SPEEDUP_FACTOR, (
        f"xl_concat string 1k: {result.cells_per_sec:.0f} cells/s, "
        f"expected >= {baseline_cps * CONCAT_1K_BASELINE_SPEEDUP_FACTOR:.0f}"
    )
