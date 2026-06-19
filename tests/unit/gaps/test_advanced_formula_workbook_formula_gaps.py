"""Formula gaps discovered via ``advanced_formula_workbook`` parity."""

from __future__ import annotations

from pathlib import Path

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.evaluator.types import XlError
from tests.unit.gaps.assertions import assert_evaluator_and_codegen_disagree
from tests.unit.gaps.workbook_helpers import (
    write_numbervalue_index_match,
    write_software_revenue_sumproduct,
    write_sumproduct_price_threshold_k24,
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


def test_software_revenue_sumproduct_returns_value_error(tmp_path: Path) -> None:
    """Category-filtered ``SUMPRODUCT`` (``K21``) returns ``#VALUE!`` in evaluator."""
    workbook = write_software_revenue_sumproduct(tmp_path / "software_revenue.xlsx")
    assert _evaluate(workbook, "Product Lookup!K21") == XlError.VALUE


def test_numbervalue_text_index_match_returns_na_string(tmp_path: Path) -> None:
    r"""``NUMBERVALUE(TEXT(INDEX(...)))`` (``K16``) yields ``"N/A"`` instead of a number."""
    workbook = write_numbervalue_index_match(tmp_path / "numbervalue_lookup.xlsx")
    assert _evaluate(workbook, "Product Lookup!K16") == "N/A"


def test_sumproduct_price_threshold_eval_codegen_mismatch(tmp_path: Path) -> None:
    """``SUMPRODUCT(($E$5:$E$19>1000)*1)`` (``K24``) disagrees evaluator ↔ codegen."""
    workbook = write_sumproduct_price_threshold_k24(tmp_path / "price_threshold_k24.xlsx")
    graph = create_dependency_graph(
        workbook,
        ["Product Lookup!K24"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    assert_evaluator_and_codegen_disagree(graph, "Product Lookup!K24")
