"""Benchmark harness for array binary-operator workloads (Sprint 0 baseline)."""

from __future__ import annotations

import json
import time
from collections.abc import Callable
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import numpy as np

from excel_grapher.core.operators import xl_concat, xl_eq, xl_gt, xl_mul
from excel_grapher.core.sumproduct import xl_sumproduct

BASELINE_VERSION = 1
DEFAULT_WARMUP_ROUNDS = 2
DEFAULT_BENCH_ROUNDS = 5

GAP_CRITERIA_SIZE = 15
MEDIUM_ARRAY_SIZE = 1_000
LARGE_ARRAY_SIZE = 10_000
SQUARE_SIDE = 100


@dataclass(frozen=True)
class OperatorWorkload:
    """One repeatable operator benchmark case."""

    name: str
    cell_count: int
    category: str
    fn: Callable[[], object]

    def run_once(self) -> object:
        return self.fn()


@dataclass(frozen=True)
class OperatorBenchResult:
    """Timing result for a single workload."""

    name: str
    cell_count: int
    category: str
    elapsed_sec: float
    cells_per_sec: float


def category_column(shape: tuple[int, ...], *, seed: int = 0) -> np.ndarray:
    """Build a Software/Hardware category column matching gap workbooks."""
    rng = np.random.default_rng(seed)
    labels = np.array(["Software", "Hardware"], dtype=object)
    flat = rng.choice(labels, size=int(np.prod(shape)))
    return flat.reshape(shape)


def numeric_column(shape: tuple[int, ...], *, seed: int = 0) -> np.ndarray:
    """Build a numeric object ndarray with Excel-like magnitudes."""
    rng = np.random.default_rng(seed)
    flat = rng.integers(50, 500, size=int(np.prod(shape)), dtype=np.int64)
    return flat.astype(object).reshape(shape)


def numeric_string_column(shape: tuple[int, ...], *, seed: int = 0) -> np.ndarray:
    """Build a column of numeric strings stored as Excel text cells."""
    rng = np.random.default_rng(seed)
    flat = rng.integers(50, 500, size=int(np.prod(shape)), dtype=np.int64)
    return flat.astype(str).astype(object).reshape(shape)


def whitespace_numeric_string_column(shape: tuple[int, ...], *, seed: int = 0) -> np.ndarray:
    """Build numeric strings with leading/trailing whitespace."""
    rng = np.random.default_rng(seed)
    flat = rng.integers(50, 500, size=int(np.prod(shape)), dtype=np.int64)
    return np.array([f" {value} " for value in flat], dtype=object).reshape(shape)


def string_prefix_column(shape: tuple[int, ...], *, seed: int = 0) -> np.ndarray:
    """Build single-letter prefixes for concat benchmarks."""
    rng = np.random.default_rng(seed)
    flat = rng.choice(np.array(list("abcdef"), dtype=object), size=int(np.prod(shape)))
    return flat.reshape(shape)


def digit_suffix_column(shape: tuple[int, ...], *, seed: int = 0) -> np.ndarray:
    """Build numeric suffixes for concat benchmarks."""
    rng = np.random.default_rng(seed)
    flat = rng.integers(0, 10, size=int(np.prod(shape)), dtype=np.int64)
    return flat.astype(object).reshape(shape)


def _shape_1d(n: int) -> tuple[int, int]:
    return (n, 1)


def _shape_square(side: int) -> tuple[int, int]:
    return (side, side)


def build_workloads() -> tuple[OperatorWorkload, ...]:
    """Return the standard Sprint 0 operator benchmark matrix."""
    gap_shape = _shape_1d(GAP_CRITERIA_SIZE)
    medium_shape = _shape_1d(MEDIUM_ARRAY_SIZE)
    large_shape = _shape_1d(LARGE_ARRAY_SIZE)
    square_shape = _shape_square(SQUARE_SIDE)

    gap_categories = category_column(gap_shape, seed=1)
    medium_categories = category_column(medium_shape, seed=2)
    large_categories = category_column(large_shape, seed=3)
    square_categories = category_column(square_shape, seed=4)
    medium_values = numeric_column(medium_shape, seed=5)
    medium_numbers = numeric_column(medium_shape, seed=7)
    large_numeric_strings = numeric_string_column(large_shape, seed=10)
    large_numeric_string_values = numeric_column(large_shape, seed=11)
    large_whitespace_numeric_strings = whitespace_numeric_string_column(large_shape, seed=12)
    medium_prefixes = string_prefix_column(medium_shape, seed=8)
    medium_suffixes = digit_suffix_column(medium_shape, seed=9)

    return (
        OperatorWorkload(
            name="xl_eq_string_gap_15",
            cell_count=GAP_CRITERIA_SIZE,
            category="compare",
            fn=lambda: xl_eq(gap_categories, "Software"),
        ),
        OperatorWorkload(
            name="xl_eq_string_1k",
            cell_count=MEDIUM_ARRAY_SIZE,
            category="compare",
            fn=lambda: xl_eq(medium_categories, "Software"),
        ),
        OperatorWorkload(
            name="xl_eq_string_10k",
            cell_count=LARGE_ARRAY_SIZE,
            category="compare",
            fn=lambda: xl_eq(large_categories, "Software"),
        ),
        OperatorWorkload(
            name="xl_gt_numeric_1k",
            cell_count=MEDIUM_ARRAY_SIZE,
            category="compare",
            fn=lambda: xl_gt(medium_numbers, 200),
        ),
        OperatorWorkload(
            name="xl_eq_numeric_string_10k",
            cell_count=LARGE_ARRAY_SIZE,
            category="compare",
            fn=lambda: xl_eq(large_numeric_strings, large_numeric_string_values),
        ),
        OperatorWorkload(
            name="xl_eq_numeric_string_ws_10k",
            cell_count=LARGE_ARRAY_SIZE,
            category="compare",
            fn=lambda: xl_eq(large_whitespace_numeric_strings, large_numeric_string_values),
        ),
        OperatorWorkload(
            name="xl_mul_numeric_1k",
            cell_count=MEDIUM_ARRAY_SIZE,
            category="arithmetic",
            fn=lambda: xl_mul(medium_values, 2.0),
        ),
        OperatorWorkload(
            name="xl_concat_string_1k",
            cell_count=MEDIUM_ARRAY_SIZE,
            category="concat",
            fn=lambda: xl_concat(medium_prefixes, medium_suffixes),
        ),
        OperatorWorkload(
            name="sumproduct_criteria_chain_1k",
            cell_count=MEDIUM_ARRAY_SIZE,
            category="integration",
            fn=lambda: xl_sumproduct(xl_mul(xl_eq(medium_categories, "Software"), medium_values)),
        ),
        OperatorWorkload(
            name="xl_eq_string_10k_square",
            cell_count=SQUARE_SIDE * SQUARE_SIDE,
            category="compare",
            fn=lambda: xl_eq(square_categories, "Software"),
        ),
    )


def bench_workload(
    workload: OperatorWorkload,
    *,
    warmup_rounds: int = DEFAULT_WARMUP_ROUNDS,
    bench_rounds: int = DEFAULT_BENCH_ROUNDS,
) -> OperatorBenchResult:
    """Time one workload and return cells/sec using the median elapsed sample."""
    for _ in range(warmup_rounds):
        workload.run_once()

    samples: list[float] = []
    for _ in range(bench_rounds):
        start = time.perf_counter()
        workload.run_once()
        samples.append(time.perf_counter() - start)

    elapsed = float(np.median(samples))
    cells_per_sec = workload.cell_count / elapsed if elapsed > 0 else float("inf")
    return OperatorBenchResult(
        name=workload.name,
        cell_count=workload.cell_count,
        category=workload.category,
        elapsed_sec=elapsed,
        cells_per_sec=cells_per_sec,
    )


def collect_baseline(
    *,
    warmup_rounds: int = DEFAULT_WARMUP_ROUNDS,
    bench_rounds: int = DEFAULT_BENCH_ROUNDS,
) -> list[OperatorBenchResult]:
    """Run all standard workloads and return timing results."""
    results: list[OperatorBenchResult] = []
    for workload in build_workloads():
        results.append(
            bench_workload(
                workload,
                warmup_rounds=warmup_rounds,
                bench_rounds=bench_rounds,
            )
        )
    return results


def baseline_document(results: list[OperatorBenchResult]) -> dict[str, Any]:
    """Serialize benchmark results into the checked-in baseline schema."""
    return {
        "version": BASELINE_VERSION,
        "workloads": [
            {
                "name": result.name,
                "cell_count": result.cell_count,
                "category": result.category,
                "elapsed_sec": round(result.elapsed_sec, 6),
                "cells_per_sec": round(result.cells_per_sec, 2),
            }
            for result in results
        ],
    }


def load_baseline_document(path: Path) -> dict[str, Any]:
    """Load a baseline JSON document from disk."""
    return json.loads(path.read_text(encoding="utf-8"))


def write_baseline_document(path: Path, results: list[OperatorBenchResult]) -> None:
    """Write benchmark results to a baseline JSON file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(baseline_document(results), indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def results_to_rows(results: list[OperatorBenchResult]) -> list[dict[str, Any]]:
    """Convert results to plain dict rows (for CLI printing)."""
    return [asdict(result) for result in results]
