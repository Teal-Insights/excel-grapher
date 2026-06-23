"""Top-level array formula results (issue #284).

Product contract: standalone formulas whose top-level result is a multi-cell
binary operation (compare, arithmetic, concat) return an ``object``-dtype
``numpy.ndarray`` at the formula anchor cell. Only 1×1 ranges auto-resolve to
scalars. Physical spill into neighbor cells is not modeled.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import numpy as np
import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.array_results import array_values_equal, is_array_result
from excel_grapher.evaluator.types import XlError
from tests.fixtures.array_results.workbook import (
    build_column_compare_workbook,
    build_grid_2d_compare_workbook,
    build_numeric_compare_workbook,
    build_row_compare_workbook,
    column_compare_path,
    ensure_committed_fixtures,
    grid_2d_compare_path,
    numeric_compare_path,
    row_compare_path,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator
from tests.unit.gaps.workbook_helpers import write_sumproduct_category_filter

ensure_committed_fixtures()

COLUMN_COMPARE_TARGET = "Data!D10"
ROW_COMPARE_TARGET = "Data!F5"
NUMERIC_COMPARE_TARGET = "Data!D10"


def _as_ndarray(value: object) -> np.ndarray:
    assert isinstance(value, np.ndarray)
    return cast(np.ndarray, value)


def _array_tolist(value: object) -> Any:
    return cast(Any, _as_ndarray(value)).tolist()


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


def _graph_from_workbook(path: Path, target: str) -> DependencyGraph:
    return create_dependency_graph(
        path,
        [target],
        load_values=True,
        use_cached_dynamic_refs=True,
    )


def _evaluate_graph(graph: DependencyGraph, target: str) -> object:
    with FormulaEvaluator(graph) as evaluator:
        return evaluator.evaluate(target)


def _evaluate_workbook(path: Path, target: str) -> object:
    graph = _graph_from_workbook(path, target)
    return _evaluate_graph(graph, target)


def _graph_with_array_compare_and(formula_address: str, formula: str) -> DependencyGraph:
    graph = DependencyGraph()
    for row, category in enumerate(["Software", "Hardware", "Software"], start=5):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node(formula_address, formula, None))
    return graph


def _assert_bool_column_array(result: object, expected: list[list[bool]]) -> np.ndarray:
    assert isinstance(result, np.ndarray)
    arr = cast(np.ndarray, result)
    assert arr.dtype == object
    assert arr.shape == (len(expected), 1)
    assert arr.tolist() == expected
    for value in arr.ravel():
        assert isinstance(value, bool)
    return arr


def _assert_bool_row_array(result: object, expected: list[list[bool]]) -> np.ndarray:
    assert isinstance(result, np.ndarray)
    arr = cast(np.ndarray, result)
    assert arr.dtype == object
    assert arr.shape == (1, len(expected[0]))
    assert arr.tolist() == expected
    for value in arr.ravel():
        assert isinstance(value, bool)
    return arr


def test_column_compare_returns_bool_ndarray() -> None:
    """``=C5:C7="Software"`` at top level is a column of booleans."""
    result = _evaluate_workbook(column_compare_path(), COLUMN_COMPARE_TARGET)
    _assert_bool_column_array(result, [[True], [False], [True]])


def test_row_compare_returns_bool_ndarray() -> None:
    """``=C5:E5="A"`` at top level is a row of booleans."""
    result = _evaluate_workbook(row_compare_path(), ROW_COMPARE_TARGET)
    _assert_bool_row_array(result, [[True, False, True]])


def test_numeric_compare_returns_bool_ndarray() -> None:
    """``=C5:C7>E5:E7`` at top level is a column of booleans."""
    result = _evaluate_workbook(numeric_compare_path(), NUMERIC_COMPARE_TARGET)
    _assert_bool_column_array(result, [[True], [False], [False]])


def test_single_cell_compare_stays_scalar() -> None:
    """1×1 range compares auto-resolve to a scalar at top level."""
    graph = DependencyGraph()
    graph.add_node(_make_node("Data!C5", None, "Software"))
    graph.add_node(_make_node("Data!D10", '=Data!C5="Software"', None))
    with FormulaEvaluator(graph) as evaluator:
        result = evaluator.evaluate("Data!D10")
    assert isinstance(result, bool)
    assert result is True


def test_sumproduct_sibling_still_returns_scalar(tmp_path: Path) -> None:
    """Aggregators keep scalar top-level results; arrays stay internal (#267)."""
    workbook = write_sumproduct_category_filter(tmp_path / "sumproduct_sibling.xlsx")
    result = _evaluate_workbook(workbook, "Product Lookup!I14")
    assert not isinstance(result, np.ndarray)
    assert result == pytest.approx(630.0)


def test_multicell_ndarray_survives_cache_round_trip() -> None:
    """Multi-cell top-level results are cached as ndarrays and reused."""
    graph = DependencyGraph()
    for row, category in enumerate(["Software", "Hardware", "Software"], start=5):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    with FormulaEvaluator(graph) as evaluator:
        first = evaluator.evaluate("Data!D10")
        second = evaluator.evaluate("Data!D10")
    assert is_array_result(first)
    assert first is second
    assert array_values_equal(first, second)


def test_array_cell_ref_binary_multiply_broadcasts() -> None:
    """``=D10*1`` element-wise coerces a cached bool column to numbers."""
    graph = _graph_with_array_compare_and("Data!E10", "=Data!D10*1")
    result = _evaluate_graph(graph, "Data!E10")
    assert isinstance(result, np.ndarray)
    assert _array_tolist(result) == [[1.0], [0.0], [1.0]]


def test_array_cell_ref_sumproduct_with_range() -> None:
    """``SUMPRODUCT`` accepts an array-result cell ref when shapes align."""
    graph = DependencyGraph()
    for row, (category, value) in enumerate(
        zip(["Software", "Hardware", "Software"], [10.0, 20.0, 30.0], strict=True),
        start=5,
    ):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
        graph.add_node(_make_node(f"Data!E{row}", None, value))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node("Data!F10", "=SUMPRODUCT(Data!D10,Data!E5:E7)", None))
    result = _evaluate_graph(graph, "Data!F10")
    assert result == pytest.approx(40.0)


def test_if_on_array_cell_returns_value_error() -> None:
    """``IF`` cannot coerce a multi-cell array operand to a scalar condition."""
    graph = _graph_with_array_compare_and("Data!E10", "=IF(Data!D10,1,0)")
    result = _evaluate_graph(graph, "Data!E10")
    assert result is XlError.VALUE


def test_grid_2d_compare_returns_bool_ndarray() -> None:
    """``=C5:D6>5`` at top level is a 2×2 boolean array."""
    result = _evaluate_workbook(grid_2d_compare_path(), "Data!D10")
    assert isinstance(result, np.ndarray)
    arr = cast(np.ndarray, result)
    assert arr.shape == (2, 2)
    assert cast(Any, arr).tolist() == [[True, False], [True, True]]


def test_spill_blocked_when_spill_slot_has_leaf_value() -> None:
    """Occupied spill slot returns ``#SPILL!`` instead of spilling the array."""
    graph = DependencyGraph()
    for row, category in enumerate(["Software", "Hardware", "Software"], start=5):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node("Data!D11", None, 1))
    result = _evaluate_graph(graph, "Data!D10")
    assert result == XlError.SPILL


def test_sum_spill_range_without_neighbor_nodes() -> None:
    """``SUM(D10:D12)`` reads virtual spill scalars from the anchor array."""
    graph = DependencyGraph()
    for row, category in enumerate(["Software", "Hardware", "Software"], start=5):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node("Data!F10", "=SUM(Data!D10:Data!D12)", None))
    result = _evaluate_graph(graph, "Data!F10")
    assert result == pytest.approx(2.0)


def test_sumproduct_spill_range_matches_cell_ref() -> None:
    """``SUMPRODUCT(D10:D12,…)`` agrees with using the anchor array cell ref."""
    graph = DependencyGraph()
    for row, (category, value) in enumerate(
        zip(["Software", "Hardware", "Software"], [10.0, 20.0, 30.0], strict=True),
        start=5,
    ):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
        graph.add_node(_make_node(f"Data!E{row}", None, value))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node("Data!F10", "=SUMPRODUCT(Data!D10,Data!E5:E7)", None))
    graph.add_node(_make_node("Data!G10", "=SUMPRODUCT(Data!D10:Data!D12,Data!E5:E7)", None))
    ref_result = _evaluate_graph(graph, "Data!F10")
    spill_result = _evaluate_graph(graph, "Data!G10")
    assert ref_result == pytest.approx(40.0)
    assert spill_result == ref_result


def test_explicit_leaf_wins_over_virtual_spill() -> None:
    """A leaf in a spill slot blocks the anchor formula with ``#SPILL!``."""
    graph = DependencyGraph()
    for row, category in enumerate(["Software", "Hardware", "Software"], start=5):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node("Data!D11", None, 0))
    assert _evaluate_graph(graph, "Data!D11") == 0
    assert _evaluate_graph(graph, "Data!D10") is XlError.SPILL


@pytest.mark.parametrize(
    ("formula_address", "formula"),
    [
        ("Data!E10", "=Data!D10*1"),
        ("Data!F10", "=SUMPRODUCT(Data!D10,Data!E5:E7)"),
        ("Data!G10", "=SUM(Data!D10:Data!D12)"),
        ("Data!H10", "=SUMPRODUCT(Data!D10:Data!D12,Data!E5:E7)"),
        ("Data!D10", '=Data!C5:C7="Software"'),
    ],
)
def test_array_consumer_eval_codegen_parity(formula_address: str, formula: str) -> None:
    """Dependent formulas using array anchors match between evaluator and export."""
    graph = DependencyGraph()
    for row, (category, value) in enumerate(
        zip(["Software", "Hardware", "Software"], [10.0, 20.0, 30.0], strict=True),
        start=5,
    ):
        graph.add_node(_make_node(f"Data!C{row}", None, category))
        graph.add_node(_make_node(f"Data!E{row}", None, value))
    graph.add_node(_make_node("Data!D10", '=Data!C5:C7="Software"', None))
    graph.add_node(_make_node(formula_address, formula, None))
    assert_codegen_matches_evaluator(graph, [formula_address])


def test_column_compare_eval_codegen_parity() -> None:
    """Evaluator and export must agree on ndarray top-level results (Sprint 2)."""
    graph = _graph_from_workbook(column_compare_path(), COLUMN_COMPARE_TARGET)
    result = assert_codegen_matches_evaluator(graph, [COLUMN_COMPARE_TARGET])
    eval_value = result.evaluator_results[COLUMN_COMPARE_TARGET]
    gen_value = result.generated_results[COLUMN_COMPARE_TARGET]
    assert isinstance(eval_value, np.ndarray)
    assert isinstance(gen_value, np.ndarray)
    assert array_values_equal(eval_value, gen_value)


def test_workbook_builders_match_committed_fixtures(tmp_path: Path) -> None:
    """Round-trip builders produce the same evaluator results as committed xlsx files."""
    cases = [
        (build_column_compare_workbook, column_compare_path(), COLUMN_COMPARE_TARGET),
        (build_row_compare_workbook, row_compare_path(), ROW_COMPARE_TARGET),
        (build_numeric_compare_workbook, numeric_compare_path(), NUMERIC_COMPARE_TARGET),
        (build_grid_2d_compare_workbook, grid_2d_compare_path(), "Data!D10"),
    ]
    for builder, committed, target in cases:
        fresh = builder(tmp_path / committed.name)
        committed_result = _evaluate_workbook(committed, target)
        fresh_result = _evaluate_workbook(fresh, target)
        if isinstance(committed_result, np.ndarray):
            assert isinstance(fresh_result, np.ndarray)
            assert array_values_equal(committed_result, fresh_result)
        else:
            assert committed_result == fresh_result
