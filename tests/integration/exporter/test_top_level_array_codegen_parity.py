"""Evaluator ↔ export parity for top-level array formula results (#284)."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import numpy as np
import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.core.array_results import array_values_equal, is_array_result
from excel_grapher.core.types import CellValue
from tests.fixtures.array_results.workbook import (
    COLUMN_COMPARE_XLSX,
    NUMERIC_COMPARE_XLSX,
    ROW_COMPARE_XLSX,
    column_compare_path,
    ensure_committed_fixtures,
    numeric_compare_path,
    row_compare_path,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator

ensure_committed_fixtures()

_ARRAY_RESULT_CASES = (
    pytest.param(column_compare_path(), "Data!D10", [[True], [False], [True]], id="column"),
    pytest.param(row_compare_path(), "Data!F5", [[True, False, True]], id="row"),
    pytest.param(numeric_compare_path(), "Data!D10", [[True], [False], [False]], id="numeric"),
)


@pytest.mark.parametrize(("workbook", "target", "expected"), _ARRAY_RESULT_CASES)
def test_top_level_array_formula_codegen_parity(
    workbook: Path,
    target: str,
    expected: list[list[bool]],
) -> None:
    """Standalone compare formulas return matching ndarrays in evaluator and export."""
    graph = create_dependency_graph(workbook, [target], load_values=True)
    result = assert_codegen_matches_evaluator(graph, [target])
    eval_value = cast(CellValue, result.evaluator_results[target])
    gen_value = cast(CellValue, result.generated_results[target])
    assert is_array_result(eval_value)
    assert is_array_result(gen_value)
    assert array_values_equal(eval_value, gen_value)
    assert cast(Any, cast(np.ndarray, eval_value)).tolist() == expected


def test_array_result_fixtures_are_committed() -> None:
    """Static workbook fixtures ship with the repo."""
    for name in (COLUMN_COMPARE_XLSX, ROW_COMPARE_XLSX, NUMERIC_COMPARE_XLSX):
        path = column_compare_path().parent / name
        assert path.is_file(), f"missing fixture: {name}"
