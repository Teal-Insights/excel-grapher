"""Dependency-graph evaluation for workbook cells stored with ``_xludf.`` prefixes."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastpyxl import load_workbook

from excel_grapher import FormulaEvaluator, create_dependency_graph
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.integration.utils.rewrite_xludf_workbook import write_xludf_workbook_copy

SANDBOX_WORKBOOK = (
    Path(__file__).resolve().parents[3] / "sandbox" / "model" / "advanced_formula_workbook.xlsx"
)

_XLUDF_TARGETS: tuple[tuple[str, str], ...] = (
    ("Product Lookup", "K7"),
    ("Product Lookup", "K9"),
    ("Product Lookup", "K12"),
    ("Formula Toolkit", "D12"),
    ("Formula Toolkit", "D30"),
)


def _skip_if_workbook_missing() -> None:
    if not SANDBOX_WORKBOOK.is_file():
        pytest.skip(f"Workbook fixture missing: {SANDBOX_WORKBOOK}")


def _xludf_formula_addresses() -> list[str]:
    return [f"'{sheet}'!{coord}" for sheet, coord in _XLUDF_TARGETS]


def test_xludf_workbook_graph_evaluator_and_codegen_parity(tmp_path: Path) -> None:
    """Evaluator and codegen agree on allowlisted ``_xludf`` cells from the sandbox workbook."""
    _skip_if_workbook_missing()
    xludf_workbook = write_xludf_workbook_copy(
        SANDBOX_WORKBOOK,
        tmp_path / "advanced_formula_workbook_xludf.xlsx",
    )
    addresses = _xludf_formula_addresses()

    wb = load_workbook(xludf_workbook, data_only=False)
    try:
        for sheet, coord in _XLUDF_TARGETS:
            formula = wb[sheet][coord].value
            assert isinstance(formula, str) and "_xludf." in formula.lower()
    finally:
        wb.close()

    graph = create_dependency_graph(
        xludf_workbook,
        addresses,
        load_values=True,
        use_cached_dynamic_refs=True,
    )

    evaluator = FormulaEvaluator(graph)
    for address in addresses:
        evaluator.evaluate(address)

    assert_codegen_matches_evaluator(graph, addresses)
