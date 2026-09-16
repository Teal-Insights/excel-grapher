"""SUMPRODUCT criteria/array semantics (issues #265 and #267).

Regression coverage for element-wise range comparisons and products inside
``SUMPRODUCT``, e.g. ``(range="label")*values`` and ``(range>threshold)*1``.
"""

# ruff: noqa: E402
from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import pytest

np = pytest.importorskip("numpy")

from excel_grapher import DependencyGraph, FormulaEvaluator, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from tests.integration.utils.parity_harness import evaluate_targets
from tests.unit.gaps.workbook_helpers import (
    write_software_revenue_sumproduct,
    write_sumproduct_category_filter,
    write_sumproduct_price_threshold_k24,
    write_sumproduct_threshold_count,
)
from tests.utils.excel_workbook_parity import assert_workbook_parity

_SUMPRODUCT_CASES: tuple[tuple[str, str, float], ...] = (
    ("cached_k21.xlsx", "Product Lookup!K21", 10598.0),
    ("cached_k24.xlsx", "Product Lookup!K24", 7.0),
    ("cached_i14.xlsx", "Product Lookup!I14", 630.0),
    ("cached_i18.xlsx", "Product Lookup!I18", 3.0),
)

_SUMPRODUCT_WRITERS = {
    "Product Lookup!K21": write_software_revenue_sumproduct,
    "Product Lookup!K24": write_sumproduct_price_threshold_k24,
    "Product Lookup!I14": write_sumproduct_category_filter,
    "Product Lookup!I18": write_sumproduct_threshold_count,
}


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


def _sumproduct_workbooks(tmp_path: Path) -> list[tuple[Path, str, float]]:
    cases: list[tuple[Path, str, float]] = []
    for filename, address, expected in _SUMPRODUCT_CASES:
        writer = _SUMPRODUCT_WRITERS[address]
        cases.append((writer(tmp_path / filename), address, expected))
    return cases


def test_standalone_range_comparison_returns_elementwise_array() -> None:
    """Multi-cell range compares at formula top level return element-wise arrays."""
    graph = DependencyGraph()
    for row, category in enumerate(["Software", "Hardware", "Software"], start=5):
        graph.add_node(_make_node(f"PL!C{row}", None, category))
    graph.add_node(
        _make_node("PL!D10", '=PL!C5:PL!C7="Software"', None),
    )
    with FormulaEvaluator(graph) as evaluator:
        result = evaluator.evaluate("PL!D10")
    actual = cast(Any, result).tolist() if isinstance(result, np.ndarray) else result
    assert actual == [[True], [False], [True]]


def test_sumproduct_criteria_evaluator_matches_excel_cached_values(tmp_path: Path) -> None:
    """Evaluator agrees with Excel cached values embedded in gap workbooks."""
    for workbook, address, _expected in _sumproduct_workbooks(tmp_path):
        graph = create_dependency_graph(
            workbook,
            [address],
            load_values=True,
            use_cached_dynamic_refs=True,
        )
        assert_workbook_parity(graph, [address])


def test_sumproduct_criteria_eval_codegen_parity(tmp_path: Path) -> None:
    """Evaluator and inverted-tree export agree on the four SUMPRODUCT gap workbooks."""
    for workbook, address, expected in _sumproduct_workbooks(tmp_path):
        graph = create_dependency_graph(
            workbook,
            [address],
            load_values=True,
            use_cached_dynamic_refs=True,
        )
        results = evaluate_targets(graph, [address])
        assert results[address] == pytest.approx(expected)
