"""SUMPRODUCT criteria/array semantics (issue #267).

Regression coverage for element-wise range comparisons and products inside
``SUMPRODUCT``, e.g. ``(range="label")*values`` and ``(range>threshold)*1``.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.evaluator.types import XlError
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.unit.gaps.workbook_helpers import (
    write_software_revenue_sumproduct,
    write_sumproduct_category_filter,
    write_sumproduct_threshold_count,
)


def _evaluate(path: Path, address: str) -> object:
    graph = create_dependency_graph(
        path,
        [address],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as evaluator:
        return evaluator.evaluate(address)


def test_software_revenue_sumproduct_category_filter(tmp_path: Path) -> None:
    """``SUMPRODUCT((cats="Software")*prices)`` returns filtered revenue (K21 / #267)."""
    workbook = write_software_revenue_sumproduct(tmp_path / "software_revenue.xlsx")
    result = _evaluate(workbook, "Product Lookup!K21")
    assert result != XlError.VALUE
    assert result == pytest.approx(10598.0)


def test_sumproduct_string_equality_filter(tmp_path: Path) -> None:
    """String equality criteria inside ``SUMPRODUCT`` (financial_model ``I14`` shape)."""
    workbook = write_sumproduct_category_filter(tmp_path / "category_filter.xlsx")
    result = _evaluate(workbook, "Product Lookup!I14")
    assert result != XlError.VALUE
    assert result == pytest.approx(630.0)


def test_sumproduct_numeric_threshold_count(tmp_path: Path) -> None:
    """Numeric comparison criteria ``(range>threshold)*1`` inside ``SUMPRODUCT``."""
    workbook = write_sumproduct_threshold_count(tmp_path / "threshold_count.xlsx")
    result = _evaluate(workbook, "Product Lookup!I18")
    assert result != XlError.VALUE
    assert result == pytest.approx(3.0)


def test_software_revenue_sumproduct_eval_codegen_parity(tmp_path: Path) -> None:
    """Evaluator and export agree on K21 category-filtered ``SUMPRODUCT``."""
    workbook = write_software_revenue_sumproduct(tmp_path / "software_revenue_parity.xlsx")
    graph = create_dependency_graph(
        workbook,
        ["Product Lookup!K21"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    result = assert_codegen_matches_evaluator(graph, ["Product Lookup!K21"])
    assert result.evaluator_results["Product Lookup!K21"] == pytest.approx(10598.0)
    assert result.generated_results["Product Lookup!K21"] == pytest.approx(10598.0)
