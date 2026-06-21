"""Advanced formula workbook gaps (issue #264).

Workbook-level repro for ``'Product Lookup'!K16`` in ``advanced_formula_workbook.xlsx``.
Unit coverage lives in ``tests/unit/evaluator/test_index_scalar_promotion.py``; remove this
module once issue #264 is fixed.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from tests.unit.gaps.workbook_helpers import write_numbervalue_index_match


def _evaluate(path: Path, address: str) -> object:
    graph = create_dependency_graph(
        path,
        [address],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as evaluator:
        return evaluator.evaluate(address)


def test_numbervalue_text_index_match_returns_na_string(tmp_path: Path) -> None:
    """``NUMBERVALUE(TEXT(INDEX(...)))`` returns the looked-up price (K16 / #264)."""
    workbook = write_numbervalue_index_match(tmp_path / "numbervalue_lookup.xlsx")
    result = _evaluate(workbook, "Product Lookup!K16")
    assert result == pytest.approx(1499.0)
