"""SUMPRODUCT criteria/array semantics (issues #265 and #267).

Regression coverage for element-wise range comparisons and products inside
``SUMPRODUCT``, e.g. ``(range="label")*values`` and ``(range>threshold)*1``.
"""

# ruff: noqa: E402
from __future__ import annotations

from collections.abc import Callable
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

_SUMPRODUCT_CASES: tuple[tuple[str, str, float, Callable[[Path], Path]], ...] = (
    ("cached_k21.xlsx", "Product Lookup!K21", 10598.0, write_software_revenue_sumproduct),
    ("cached_k24.xlsx", "Product Lookup!K24", 7.0, write_sumproduct_price_threshold_k24),
    ("cached_i14.xlsx", "Product Lookup!I14", 630.0, write_sumproduct_category_filter),
    ("cached_i18.xlsx", "Product Lookup!I18", 3.0, write_sumproduct_threshold_count),
)
_SUMPRODUCT_IDS = [address for _filename, address, _expected, _writer in _SUMPRODUCT_CASES]


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


def _sumproduct_graph(
    tmp_path: Path,
    filename: str,
    address: str,
    write_workbook: Callable[[Path], Path],
) -> DependencyGraph:
    workbook = write_workbook(tmp_path / filename)
    return create_dependency_graph(
        workbook,
        [address],
        load_values=True,
        use_cached_dynamic_refs=True,
    )


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


@pytest.mark.parametrize(
    ("filename", "address", "write_workbook"),
    [(filename, address, writer) for filename, address, _, writer in _SUMPRODUCT_CASES],
    ids=_SUMPRODUCT_IDS,
)
def test_sumproduct_criteria_evaluator_matches_excel_cached_values(
    tmp_path: Path,
    filename: str,
    address: str,
    write_workbook: Callable[[Path], Path],
) -> None:
    """Evaluator agrees with Excel cached values embedded in gap workbooks."""
    graph = _sumproduct_graph(tmp_path, filename, address, write_workbook)
    assert_workbook_parity(graph, [address])


@pytest.mark.parametrize(
    ("filename", "address", "expected", "write_workbook"),
    _SUMPRODUCT_CASES,
    ids=_SUMPRODUCT_IDS,
)
def test_sumproduct_criteria_eval_codegen_parity(
    tmp_path: Path,
    filename: str,
    address: str,
    expected: float,
    write_workbook: Callable[[Path], Path],
) -> None:
    """Evaluator and inverted-tree export agree on the SUMPRODUCT gap workbook."""
    graph = _sumproduct_graph(tmp_path, filename, address, write_workbook)
    results = evaluate_targets(graph, [address])
    assert results[address] == pytest.approx(expected)
