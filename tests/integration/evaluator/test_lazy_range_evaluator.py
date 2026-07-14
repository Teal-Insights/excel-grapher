"""Lazy-range selective access on FormulaEvaluator (#336 / #314).

Mirrors export lazy-range scenarios against the evaluator directly, asserting
unused sibling cells are never evaluated (via `on_cell_evaluated` / `_cache`).
"""

from __future__ import annotations

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.evaluator.types import XlError


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


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_index_over_range_with_unrelated_error_cell() -> None:
    """INDEX(A1:A5, 2) succeeds even though A4 contains a division error."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!A3", None, 30),
        _make_node("S!A4", "=1/0", None),
        _make_node("S!A5", None, 50),
        _make_node("S!B1", "=INDEX(S!A1:S!A5, 2)", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 20}


def test_index_does_not_evaluate_unused_formula_cells() -> None:
    """Evaluator INDEX leaves unused sibling formula cells unevaluated."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", "=S!A1+S!A2", None),
        _make_node("S!B1", "=INDEX(S!A1:S!A3, 2)", None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 2}
        assert "S!A3" not in ev._cache
    assert "S!A3" not in seen


def test_match_over_column_slice_of_index() -> None:
    """MATCH consumes an INDEX column view without evaluating unused cells."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", None, 2),
        _make_node("S!A2", None, 4),
        _make_node("S!B2", None, 5),
        _make_node("S!C1", "=MATCH(5, INDEX(S!A1:S!B2,,2), 0)", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": 2}


def test_match_does_not_evaluate_trailing_unused_cells() -> None:
    """Exact MATCH stops once the match is found."""
    graph = _make_graph(
        _make_node("S!A1", None, "x"),
        _make_node("S!A2", None, "y"),
        _make_node("S!A3", "=1/0", None),
        _make_node("S!B1", '=MATCH("y", S!A1:S!A3, 0)', None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 2}
        assert "S!A3" not in ev._cache
    assert "S!A3" not in seen


def test_vlookup_stops_before_trailing_unused_cells() -> None:
    """Exact VLOOKUP does not force evaluation past the matched row."""
    graph = _make_graph(
        _make_node("S!A1", None, "k1"),
        _make_node("S!B1", None, 100),
        _make_node("S!A2", None, "k2"),
        _make_node("S!B2", None, 200),
        _make_node("S!A3", "=1/0", None),
        _make_node("S!B3", None, 300),
        _make_node("S!C1", '=VLOOKUP("k1", S!A1:S!B3, 2, FALSE)', None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": 100}
        assert "S!A3" not in ev._cache
        assert "S!B3" not in ev._cache
    assert "S!A3" not in seen
    assert "S!B3" not in seen


def test_hlookup_stops_before_trailing_unused_cells() -> None:
    """Exact HLOOKUP does not force evaluation past the matched column."""
    graph = _make_graph(
        _make_node("S!A1", None, "k1"),
        _make_node("S!B1", None, "k2"),
        _make_node("S!C1", "=1/0", None),
        _make_node("S!A2", None, 100),
        _make_node("S!B2", None, 200),
        _make_node("S!C2", None, 300),
        _make_node("S!D1", '=HLOOKUP("k1", S!A1:S!C2, 2, FALSE)', None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!D1"]) == {"S!D1": 100}
        assert "S!C1" not in ev._cache
        assert "S!C2" not in ev._cache
    assert "S!C1" not in seen
    assert "S!C2" not in seen


def test_lookup_stops_before_trailing_unused_cells() -> None:
    """Vector LOOKUP does not force evaluation past the matched position."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!A3", "=1/0", None),
        _make_node("S!B1", None, "ten"),
        _make_node("S!B2", None, "twenty"),
        _make_node("S!B3", None, "thirty"),
        _make_node("S!C1", "=LOOKUP(10, S!A1:S!A3, S!B1:S!B3)", None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": "ten"}
        assert "S!A3" not in ev._cache
        assert "S!B3" not in ev._cache
    assert "S!A3" not in seen
    assert "S!B3" not in seen


def test_xlookup_stops_before_trailing_unused_cells() -> None:
    """Exact XLOOKUP does not force evaluation past the matched key."""
    graph = _make_graph(
        _make_node("S!A1", None, "k1"),
        _make_node("S!A2", None, "k2"),
        _make_node("S!A3", "=1/0", None),
        _make_node("S!B1", None, 100),
        _make_node("S!B2", None, 200),
        _make_node("S!B3", None, 300),
        _make_node("S!C1", '=_xlfn.XLOOKUP("k1", S!A1:S!A3, S!B1:S!B3)', None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": 100}
        assert "S!A3" not in ev._cache
        assert "S!B3" not in ev._cache
    assert "S!A3" not in seen
    assert "S!B3" not in seen


def test_sum_over_range_with_error_produces_error_code() -> None:
    """SUM over a range containing an error surfaces the first error (reductions stay eager)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", "=1/0", None),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", "=SUM(S!A1:S!A3)", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": XlError.DIV}


def test_sum_still_evaluates_all_cells_in_range() -> None:
    """Full-scan reductions still visit every cell (eager until Phase 3)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", "=S!A1+1", None),
        _make_node("S!A3", "=S!A1+2", None),
        _make_node("S!B1", "=SUM(S!A1:S!A3)", None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 6}
    assert "S!A1" in seen
    assert "S!A2" in seen
    assert "S!A3" in seen


def test_match_over_offset_does_not_evaluate_trailing_unused_cells() -> None:
    """OFFSET → MATCH stays selective under lazy-by-default range resolution."""
    graph = _make_graph(
        _make_node("S!A1", None, "x"),
        _make_node("S!A2", None, "y"),
        _make_node("S!A3", "=1/0", None),
        _make_node("S!B1", '=MATCH("y", OFFSET(S!A1,0,0,3,1), 0)', None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 2}
        assert "S!A3" not in ev._cache
    assert "S!A3" not in seen


def test_vlookup_over_offset_stops_before_trailing_unused_cells() -> None:
    """OFFSET → VLOOKUP does not force evaluation past the matched row."""
    graph = _make_graph(
        _make_node("S!A1", None, "k1"),
        _make_node("S!B1", None, 100),
        _make_node("S!A2", None, "k2"),
        _make_node("S!B2", None, 200),
        _make_node("S!A3", "=1/0", None),
        _make_node("S!B3", None, 300),
        _make_node(
            "S!C1",
            '=VLOOKUP("k1", OFFSET(S!A1,0,0,3,2), 2, FALSE)',
            None,
        ),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": 100}
        assert "S!A3" not in ev._cache
        assert "S!B3" not in ev._cache
    assert "S!A3" not in seen
    assert "S!B3" not in seen


def test_binary_op_over_ranges_does_not_eager_resolve_excel_range(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Binary range ops bind lazy Range; resolve_excel_range is never called."""
    from excel_grapher.core import types as core_types

    def _boom(*_args: object, **_kwargs: object) -> object:
        raise AssertionError("resolve_excel_range must not run for binary operands")

    monkeypatch.setattr(core_types, "resolve_excel_range", _boom)
    monkeypatch.setattr(
        "excel_grapher.evaluator.evaluator.resolve_excel_range",
        _boom,
    )

    graph = _make_graph(
        _make_node("S!A1", None, "x"),
        _make_node("S!A2", None, "y"),
        _make_node("S!A3", None, "x"),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!B3", None, 30),
        _make_node("S!C1", '=SUMPRODUCT((S!A1:S!A3="x")*S!B1:S!B3)', None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": 40.0}


def test_array_unary_negation_over_range() -> None:
    """Unary minus over a multi-cell range maps element-wise (export parity)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", "=SUM(-(S!A1:S!A3))", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": -6.0}


def test_criteria_product_sum_parity_via_binary_ops() -> None:
    """Criteria comparison * values reduces correctly via lazy Grid maps."""
    graph = _make_graph(
        _make_node("S!A1", None, "x"),
        _make_node("S!A2", None, "y"),
        _make_node("S!A3", None, "x"),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!B3", None, 30),
        _make_node("S!C1", '=SUM((S!A1:S!A3="x")*S!B1:S!B3)', None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": 40.0}


def test_binary_op_fail_fast_does_not_evaluate_trailing_formula_cells() -> None:
    """Array multiply stops at the first cell error; trailing formulas stay cold."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", "=1/0", None),
        _make_node("S!A3", "=S!A1+99", None),
        _make_node("S!B1", None, 2),
        _make_node("S!B2", None, 3),
        _make_node("S!B3", "=S!B1+99", None),
        _make_node("S!C1", "=S!A1:S!A3*S!B1:S!B3", None),
    )
    seen: list[str] = []

    def _track(address: str, _value: object) -> None:
        seen.append(address)

    with FormulaEvaluator(graph, on_cell_evaluated=_track) as ev:
        assert ev.evaluate(["S!C1"]) == {"S!C1": XlError.DIV}
        assert "S!A3" not in ev._cache
        assert "S!B3" not in ev._cache
    assert "S!A3" not in seen
    assert "S!B3" not in seen
