"""Vectorized numeric arithmetic fast path for binary operators."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import numpy as np
import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.operators import xl_add, xl_div, xl_mul, xl_pow, xl_sub
from excel_grapher.core.operators_bench import (
    bench_workload,
    build_workloads,
    load_baseline_document,
)
from excel_grapher.core.operators_reference import broadcast_pair, reference_arithmetic_array
from excel_grapher.core.types import CellValue, XlError
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.unit.core.test_operators_baseline import BASELINE_PATH
from tests.unit.gaps.workbook_helpers import write_large_numeric_sumproduct

LARGE_SHAPE = (2_000, 1)
MUL_1K_BASELINE_SPEEDUP_FACTOR = 5.0


def _as_ndarray(value: object) -> np.ndarray:
    assert isinstance(value, np.ndarray)
    return cast(np.ndarray, value)


def _assert_matches_reference(op: str, left: CellValue, right: CellValue) -> None:
    pair = broadcast_pair(left, right)
    assert not isinstance(pair, XlError)
    arr_left, arr_right = pair
    expected = reference_arithmetic_array(op, arr_left, arr_right)
    dispatch = {
        "+": xl_add,
        "-": xl_sub,
        "*": xl_mul,
        "/": xl_div,
        "^": xl_pow,
    }
    actual = dispatch[op](left, right)
    if isinstance(expected, XlError):
        assert actual == expected
        return
    assert isinstance(actual, np.ndarray)
    assert cast(Any, actual).tolist() == cast(Any, expected).tolist()


def _numeric_array(shape: tuple[int, ...], *, seed: int) -> np.ndarray:
    rng = np.random.default_rng(seed)
    flat = rng.integers(1, 50, size=int(np.prod(shape)), dtype=np.int64)
    return flat.astype(object).reshape(shape)


@pytest.mark.parametrize("op", ["+", "-", "*", "/", "^"])
def test_numeric_fastpath_matches_reference_on_large_arrays(op: str) -> None:
    left = _numeric_array(LARGE_SHAPE, seed=11)
    right = _numeric_array(LARGE_SHAPE, seed=12)
    _assert_matches_reference(op, left, right)


@pytest.mark.parametrize("op", ["+", "-", "*", "/", "^"])
def test_numeric_fastpath_matches_reference_with_scalar_broadcast(op: str) -> None:
    left = _numeric_array((500, 1), seed=21)
    _assert_matches_reference(op, left, 2.0)


def test_numeric_fastpath_matches_reference_with_bools_and_none() -> None:
    left = np.array([[True, None], [False, 3.0]], dtype=object)
    right = np.array([[2.0, 4.0], [1.0, 2.0]], dtype=object)
    _assert_matches_reference("*", left, right)


def test_numeric_fastpath_matches_reference_with_numeric_strings() -> None:
    left = np.array([["10", " 2.5 "], ["0", ""]], dtype=object)
    right = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    _assert_matches_reference("+", left, right)


def test_numeric_fastpath_falls_back_on_embedded_error() -> None:
    left = np.array([[1.0, XlError.NA], [3.0, 4.0]], dtype=object)
    right = np.array([[2.0, 2.0], [2.0, 2.0]], dtype=object)
    assert xl_mul(left, right) == XlError.NA


def test_numeric_fastpath_falls_back_on_non_numeric_string() -> None:
    left = np.array([["abc", 2.0]], dtype=object)
    right = np.array([[1.0, 2.0]], dtype=object)
    assert xl_mul(left, right) == XlError.VALUE


def test_numeric_fastpath_div_fail_fast_on_first_zero_in_c_order() -> None:
    left = np.array([[1.0, 2.0], [3.0, 4.0]], dtype=object)
    right = np.array([[1.0, 0.0], [2.0, 4.0]], dtype=object)
    assert xl_div(left, right) == XlError.DIV


def test_numeric_fastpath_pow_invalid_returns_num() -> None:
    left = np.array([[-1.0, 2.0]], dtype=object)
    right = np.array([[0.5, 2.0]], dtype=object)
    assert xl_pow(left, right) == XlError.NUM


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def test_large_in_memory_range_mul_matches_reference() -> None:
    """500-cell range multiply stays aligned with the reference loop."""
    graph = DependencyGraph()
    for row in range(1, 501):
        graph.add_node(_make_node(f"S!A{row}", None, float(row % 7 + 1)))
        graph.add_node(_make_node(f"S!B{row}", None, 3.0))
    graph.add_node(_make_node("S!C1", "=S!A1:A500*S!B1:B500", None))
    with FormulaEvaluator(graph) as evaluator:
        result = evaluator.evaluate("S!C1")
    left = np.array([[float(i % 7 + 1)] for i in range(1, 501)], dtype=object)
    right = np.array([[3.0] for _ in range(500)], dtype=object)
    expected = reference_arithmetic_array("*", left, right)
    assert isinstance(result, np.ndarray)
    assert cast(Any, result).tolist() == cast(Any, expected).tolist()


def test_large_numeric_sumproduct_eval_codegen_parity(tmp_path: Path) -> None:
    """Evaluator and export agree on large ``SUMPRODUCT`` with numeric range product."""
    workbook = write_large_numeric_sumproduct(tmp_path / "large_sumproduct.xlsx", rows=500)
    graph = create_dependency_graph(
        workbook,
        ["Data!C1"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    result = assert_codegen_matches_evaluator(graph, ["Data!C1"])
    assert result.evaluator_results["Data!C1"] == pytest.approx(15_000.0)
    assert result.generated_results["Data!C1"] == pytest.approx(15_000.0)


@pytest.mark.slow
def test_xl_mul_numeric_1k_beats_baseline() -> None:
    """Numeric multiply fast path should materially exceed the loop baseline throughput."""
    workload = next(w for w in build_workloads() if w.name == "xl_mul_numeric_1k")
    baseline_doc = load_baseline_document(BASELINE_PATH)["workloads"]
    baseline_cps = next(
        entry["cells_per_sec"] for entry in baseline_doc if entry["name"] == "xl_mul_numeric_1k"
    )
    result = bench_workload(workload, warmup_rounds=2, bench_rounds=5)
    assert result.cells_per_sec >= baseline_cps * MUL_1K_BASELINE_SPEEDUP_FACTOR, (
        f"xl_mul 1k: {result.cells_per_sec:.0f} cells/s, "
        f"expected >= {baseline_cps * MUL_1K_BASELINE_SPEEDUP_FACTOR:.0f}"
    )
