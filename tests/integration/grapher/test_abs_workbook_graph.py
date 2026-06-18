"""Workbook-backed ABS evaluation for ``normdist_sigma_band`` formulas."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator

SANDBOX_WORKBOOK = (
    Path(__file__).resolve().parents[3] / "sandbox" / "model" / "advanced_formula_workbook.xlsx"
)

_SIGMA_BAND_TARGETS: tuple[str, ...] = tuple(
    f"'Statistical Analysis'!M{row}" for row in range(19, 23)
)


def test_normdist_sigma_band_workbook_graph_parity() -> None:
    """``normdist_sigma_band`` cells with nested ``ABS`` evaluate and codegen cleanly."""
    if not SANDBOX_WORKBOOK.is_file():
        pytest.skip(f"Workbook fixture missing: {SANDBOX_WORKBOOK}")

    graph = create_dependency_graph(
        SANDBOX_WORKBOOK,
        list(_SIGMA_BAND_TARGETS),
        load_values=True,
        use_cached_dynamic_refs=True,
    )

    evaluator = FormulaEvaluator(graph)
    for address in _SIGMA_BAND_TARGETS:
        value = evaluator.evaluate(address)
        assert isinstance(value, str)
        assert value in {"Within 1σ", "Within 2σ", "Outlier >2σ"}

    assert_codegen_matches_evaluator(graph, list(_SIGMA_BAND_TARGETS))
