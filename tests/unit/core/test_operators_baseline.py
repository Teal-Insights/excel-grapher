"""Baseline semantics and throughput for vectorized binary-operator work.

Records the per-cell loop semantics contract, reference-path equivalence,
and checked-in throughput numbers that later sprints must beat while preserving
Excel coercion and error behavior.
Semantics contract (array paths):

- **Comparisons**: fail-fast in C-order — the first embedded ``XlError`` during
  row-major ``np.ndindex`` iteration wins; no partial result array is returned.
  Numeric coercion via ``to_number``; when either side fails, fall back to
  casefolded string comparison of ``to_string`` values.
- **Arithmetic**: preserve per-element errors in the result array (operand
  sentinels, coercion failures, ``/`` ``DIV`` on zero divisor, ``^`` ``NUM`` on
  invalid or complex results). Top-level scalar operand errors still propagate.
- **Concat**: per-cell ``to_string``; top-level operand errors propagate before
  the array loop.
- **Broadcasting**: shape mismatch returns ``VALUE``; scalars broadcast with
  ``np.full(..., dtype=object)``.

Fast-path work should target **>=10x** ``cells_per_sec`` on 10K-cell compare and
multiply workloads while keeping gap-sized (~15 cell) paths within ~5% of this
baseline.
"""

# ruff: noqa: E402
from __future__ import annotations

import pytest

np = pytest.importorskip("numpy")

from excel_grapher.core.operators import xl_concat, xl_div, xl_eq, xl_mul
from excel_grapher.core.types import XlError
from tests.bench.operators_bench import (
    BASELINE_VERSION,
    build_workloads,
    collect_baseline,
    load_baseline_document,
)
from tests.paths import OPERATORS_BASELINE_FIXTURES
from tests.unit.core.operators_test_helpers import (
    ARITHMETIC_DISPATCH,
    COMPARE_DISPATCH,
    COMPARE_OPS,
    array_tolist,
    assert_cellvalue_equal,
    reference_arithmetic,
    reference_compare,
    reference_concat,
)

BASELINE_PATH = OPERATORS_BASELINE_FIXTURES / "baseline.json"

EXPECTED_WORKLOAD_NAMES = frozenset(
    {
        "xl_eq_string_gap_15",
        "xl_eq_string_1k",
        "xl_eq_string_10k",
        "xl_eq_numeric_string_10k",
        "xl_eq_numeric_string_ws_10k",
        "xl_gt_numeric_1k",
        "xl_mul_numeric_1k",
        "xl_concat_string_1k",
        "sumproduct_criteria_chain_1k",
        "xl_eq_string_10k_square",
    }
)


@pytest.mark.parametrize("op", COMPARE_OPS)
def test_reference_compare_matches_public_compare(op: str) -> None:
    left = np.array([["Software", 2.0], [XlError.NA, "Hardware"]], dtype=object)
    right = np.array([["software", "2"], [1.0, "hardware"]], dtype=object)
    assert_cellvalue_equal(
        COMPARE_DISPATCH[op](left, right),
        reference_compare(op, left, right),
    )


@pytest.mark.parametrize(
    ("op", "dispatch"),
    [(op, fn) for op, fn in ARITHMETIC_DISPATCH.items()],
)
def test_reference_arithmetic_matches_public_arithmetic(op: str, dispatch) -> None:
    left = np.array([[2.0, 4.0], [9.0, -1.0]], dtype=object)
    right = np.array([[1.0, 2.0], [0.5, 2.0]], dtype=object)
    assert_cellvalue_equal(dispatch(left, right), reference_arithmetic(op, left, right))


def test_reference_concat_matches_public_concat() -> None:
    left = np.array([["a", 1.0], [True, None]], dtype=object)
    right = np.array([["z", 2], ["!", ""]], dtype=object)
    assert_cellvalue_equal(xl_concat(left, right), reference_concat(left, right))


def test_array_compare_fail_fast_returns_first_error_in_c_order() -> None:
    left = np.array([[1.0, XlError.DIV], [XlError.NA, 3.0]], dtype=object)
    assert xl_eq(left, 1.0) == XlError.DIV


def test_array_arithmetic_preserves_embedded_errors_per_element() -> None:
    left = np.array([[1.0, XlError.REF], [XlError.NA, 3.0]], dtype=object)
    assert array_tolist(xl_mul(left, 2.0)) == [
        [2.0, XlError.REF],
        [XlError.NA, 6.0],
    ]


def test_array_division_preserves_div_zero_per_element() -> None:
    left = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    right = np.array([[1.0, 0.0], [2.0, 4.0]], dtype=object)
    assert array_tolist(xl_div(left, right)) == [
        [1.0, XlError.DIV],
        [1.5, 1.0],
    ]


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
    assert array_tolist(xl_eq(left, right)) == [[True, True]]


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
