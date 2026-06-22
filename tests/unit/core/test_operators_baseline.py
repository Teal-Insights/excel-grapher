"""Baseline semantics and throughput for vectorized binary-operator work.

Records the per-cell loop semantics contract, reference-path equivalence,
and checked-in throughput numbers that later sprints must beat while preserving
Excel coercion and error behavior.

Semantics contract (array paths):

- **Fail-fast, C-order**: the first embedded ``XlError`` during row-major
  ``np.ndindex`` iteration wins; no partial result array is returned.
- **Comparisons**: numeric coercion via ``to_number``; when either side fails,
  fall back to casefolded string comparison of ``to_string`` values.
- **Arithmetic**: per-cell ``to_number``; ``/`` returns ``DIV`` on zero divisor;
  ``^`` returns ``NUM`` on invalid or complex results.
- **Concat**: per-cell ``to_string``; top-level operand errors propagate before
  the array loop.
- **Broadcasting**: shape mismatch returns ``VALUE``; scalars broadcast with
  ``np.full(..., dtype=object)``.

Fast-path work should target **>=10x** ``cells_per_sec`` on 10K-cell compare and
multiply workloads while keeping gap-sized (~15 cell) paths within ~5% of this
baseline.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import numpy as np
import pytest

from excel_grapher.core.operators import (
    xl_add,
    xl_concat,
    xl_div,
    xl_eq,
    xl_ge,
    xl_gt,
    xl_le,
    xl_lt,
    xl_mul,
    xl_ne,
    xl_pow,
    xl_sub,
)
from excel_grapher.core.operators_bench import (
    BASELINE_VERSION,
    build_workloads,
    collect_baseline,
    load_baseline_document,
)
from excel_grapher.core.operators_reference import (
    broadcast_pair,
    reference_arithmetic_array,
    reference_compare_array,
    reference_concat_array,
)
from excel_grapher.core.types import CellValue, XlError

BASELINE_PATH = (
    Path(__file__).resolve().parents[2] / "fixtures" / "operators_baseline" / "baseline.json"
)

EXPECTED_WORKLOAD_NAMES = frozenset(
    {
        "xl_eq_string_gap_15",
        "xl_eq_string_1k",
        "xl_eq_string_10k",
        "xl_gt_numeric_1k",
        "xl_mul_numeric_1k",
        "xl_concat_string_1k",
        "sumproduct_criteria_chain_1k",
        "xl_eq_string_10k_square",
    }
)

COMPARE_OPS = ("=", "<>", "<", ">", "<=", ">=")


def _assert_cellvalue_equal(actual: object, expected: object) -> None:
    if isinstance(actual, np.ndarray) and isinstance(expected, np.ndarray):
        assert actual.shape == expected.shape
        actual_list = cast(Any, actual).tolist()
        expected_list = cast(Any, expected).tolist()
        assert actual_list == expected_list
        return
    assert actual == expected


def _as_ndarray(value: object) -> np.ndarray:
    assert isinstance(value, np.ndarray)
    return cast(np.ndarray, value)


def _via_reference_compare(op: str, left: CellValue, right: CellValue) -> object:
    pair = broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    return reference_compare_array(op, pair[0], pair[1])


def _via_reference_arithmetic(op: str, left: CellValue, right: CellValue) -> object:
    pair = broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    return reference_arithmetic_array(op, pair[0], pair[1])


def _via_reference_concat(left: CellValue, right: CellValue) -> object:
    pair = broadcast_pair(left, right)
    if isinstance(pair, XlError):
        return pair
    return reference_concat_array(pair[0], pair[1])


@pytest.mark.parametrize("op", COMPARE_OPS)
def test_reference_compare_matches_public_compare(op: str) -> None:
    left = np.array([["Software", 2.0], [XlError.NA, "Hardware"]], dtype=object)
    right = np.array([["software", "2"], [1.0, "hardware"]], dtype=object)
    dispatch = {
        "=": xl_eq,
        "<>": xl_ne,
        "<": xl_lt,
        ">": xl_gt,
        "<=": xl_le,
        ">=": xl_ge,
    }
    _assert_cellvalue_equal(dispatch[op](left, right), _via_reference_compare(op, left, right))


@pytest.mark.parametrize(
    ("op", "dispatch"),
    [
        ("+", xl_add),
        ("-", xl_sub),
        ("*", xl_mul),
        ("/", xl_div),
        ("^", xl_pow),
    ],
)
def test_reference_arithmetic_matches_public_arithmetic(op: str, dispatch) -> None:
    left = np.array([[2.0, 4.0], [9.0, -1.0]], dtype=object)
    right = np.array([[1.0, 2.0], [0.5, 2.0]], dtype=object)
    _assert_cellvalue_equal(dispatch(left, right), _via_reference_arithmetic(op, left, right))


def test_reference_concat_matches_public_concat() -> None:
    left = np.array([["a", 1.0], [True, None]], dtype=object)
    right = np.array([["z", 2], ["!", ""]], dtype=object)
    _assert_cellvalue_equal(xl_concat(left, right), _via_reference_concat(left, right))


def test_array_compare_fail_fast_returns_first_error_in_c_order() -> None:
    left = np.array([[1.0, XlError.DIV], [XlError.NA, 3.0]], dtype=object)
    assert xl_eq(left, 1.0) == XlError.DIV


def test_array_arithmetic_fail_fast_returns_first_error_in_c_order() -> None:
    left = np.array([[1.0, XlError.REF], [XlError.NA, 3.0]], dtype=object)
    assert xl_mul(left, 2.0) == XlError.REF


def test_array_division_fail_fast_on_first_zero_divisor() -> None:
    left = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    right = np.array([[1.0, 0.0], [2.0, 4.0]], dtype=object)
    assert xl_div(left, right) == XlError.DIV


def test_array_shape_mismatch_returns_value_for_compare() -> None:
    left = np.array([[1.0, 2.0]], dtype=object)
    right = np.array([[1.0, 2.0, 3.0]], dtype=object)
    assert xl_eq(left, right) == XlError.VALUE


def test_top_level_error_propagates_before_array_compare() -> None:
    left = np.array([[1.0]], dtype=object)
    assert xl_eq(XlError.REF, left) == XlError.REF
    assert xl_eq(left, XlError.VALUE) == XlError.VALUE


def test_string_compare_uses_casefolded_fallback_when_numeric_coercion_fails() -> None:
    left = np.array([["TRUE", "AbC"]], dtype=object)
    right = np.array([[True, "aBc"]], dtype=object)
    assert _as_ndarray(xl_eq(left, right)).tolist() == [[True, True]]


def test_baseline_fixture_exists_and_matches_schema() -> None:
    document = load_baseline_document(BASELINE_PATH)
    assert document["version"] == BASELINE_VERSION
    workloads = document["workloads"]
    assert isinstance(workloads, list)
    assert {entry["name"] for entry in workloads} == EXPECTED_WORKLOAD_NAMES
    for entry in workloads:
        assert entry["cell_count"] > 0
        assert entry["elapsed_sec"] > 0
        assert entry["cells_per_sec"] > 0
        assert entry["category"] in {"compare", "arithmetic", "concat", "integration"}


def test_benchmark_workload_registry_matches_fixture_names() -> None:
    assert {workload.name for workload in build_workloads()} == EXPECTED_WORKLOAD_NAMES


def test_baseline_targets_include_gap_and_large_ranges() -> None:
    document = load_baseline_document(BASELINE_PATH)
    by_name = {entry["name"]: entry for entry in document["workloads"]}
    assert by_name["xl_eq_string_gap_15"]["cell_count"] == 15
    assert by_name["xl_eq_string_10k"]["cell_count"] == 10_000
    assert by_name["xl_eq_string_10k_square"]["cell_count"] == 10_000


@pytest.mark.slow
def test_refresh_operator_baseline_throughput() -> None:
    """Opt-in timing run; not part of default CI (``pytest -m slow``)."""
    results = collect_baseline(warmup_rounds=1, bench_rounds=3)
    document = load_baseline_document(BASELINE_PATH)
    recorded = {entry["name"]: entry for entry in document["workloads"]}
    for result in results:
        baseline = recorded[result.name]
        # Guard against catastrophic regressions only (order-of-magnitude).
        assert result.cells_per_sec >= baseline["cells_per_sec"] * 0.1, (
            f"{result.name}: {result.cells_per_sec:.2f} cells/s "
            f"vs baseline {baseline['cells_per_sec']:.2f}"
        )
