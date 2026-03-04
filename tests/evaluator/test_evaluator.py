import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.export_runtime.cache import CircularReferenceWarning
from excel_grapher.evaluator.name_utils import parse_address
from excel_grapher.evaluator.types import XlError


def _make_node(address: str, formula: str | None, value: object) -> Node:
    """Helper to create a Node from a sheet-qualified address."""
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
    """Helper to create a DependencyGraph from nodes."""
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_evaluator_returns_values_for_non_formula_cells() -> None:
    graph = _make_graph(_make_node("S!A1", None, 3))
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!A1"]) == {"S!A1": 3}


def test_evaluator_evaluates_cell_reference_formula() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 7),
        _make_node("S!B1", "=S!A1", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 7}


def test_evaluator_vlookup_ignores_errors_in_unrelated_table_cells() -> None:
    graph = _make_graph(
        _make_node("S!B1", None, "key"),
        _make_node("S!C1", None, 1),
        _make_node("S!B2", None, "other"),
        _make_node("S!C2", None, XlError.NA),
        _make_node("S!A1", '=VLOOKUP("key",S!B1:S!C2,2,FALSE)', None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!A1"]) == {"S!A1": 1}


def test_evaluator_on_cell_evaluated_hook_can_assert_cached_values() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", "=S!A1+1", 2),
    )
    seen: list[str] = []

    def _assert_cached(address: str, value: object) -> None:
        seen.append(address)
        node = graph.get_node(address)
        if node is None or node.formula is None or node.value is None:
            return
        assert value == node.value

    with FormulaEvaluator(graph, on_cell_evaluated=_assert_cached) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 2.0}

    assert "S!A1" in seen
    assert "S!B1" in seen


def test_evaluator_memoizes_cell_computation() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 2),
        _make_node("S!B1", "=SUM(S!A1, S!A1)", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"]) == {"S!B1": 4.0}
        # A1 should be cached after evaluation
        assert ev._cache["S!A1"] == 2  # noqa: SLF001


def test_evaluator_detects_cycles() -> None:
    graph = _make_graph(
        _make_node("S!A1", "=S!B1", None),
        _make_node("S!B1", "=S!A1", None),
    )
    with FormulaEvaluator(graph) as ev, pytest.warns(CircularReferenceWarning):
        assert ev.evaluate(["S!A1"]) == {"S!A1": 0}


def test_evaluator_raises_for_unimplemented_function() -> None:
    graph = _make_graph(_make_node("S!A1", "=no_such_function(1)", None))
    with FormulaEvaluator(graph) as ev, pytest.raises(
        NotImplementedError,
        match=r"Excel function not implemented: NO_SUCH_FUNCTION",
    ):
        ev.evaluate(["S!A1"])


# --- Operator evaluation tests ---


def test_evaluator_arithmetic_operators() -> None:
    """Test arithmetic operators: +, -, *, /, ^"""
    graph = _make_graph(
        _make_node("S!A1", "=1+2", None),
        _make_node("S!A2", "=5-3", None),
        _make_node("S!A3", "=2*3", None),
        _make_node("S!A4", "=6/2", None),
        _make_node("S!A5", "=2^3", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3", "S!A4", "S!A5"])
        assert result["S!A1"] == 3.0
        assert result["S!A2"] == 2.0
        assert result["S!A3"] == 6.0
        assert result["S!A4"] == 3.0
        assert result["S!A5"] == 8.0


def test_evaluator_comparison_operators() -> None:
    """Test comparison operators: <, >, =, <=, >=, <>"""
    graph = _make_graph(
        _make_node("S!A1", "=1<2", None),
        _make_node("S!A2", "=2>1", None),
        _make_node("S!A3", "=1=1", None),
        _make_node("S!A4", "=1<=1", None),
        _make_node("S!A5", "=2>=1", None),
        _make_node("S!A6", "=1<>2", None),
        _make_node("S!A7", "=1<>1", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3", "S!A4", "S!A5", "S!A6", "S!A7"])
        assert result["S!A1"] is True
        assert result["S!A2"] is True
        assert result["S!A3"] is True
        assert result["S!A4"] is True
        assert result["S!A5"] is True
        assert result["S!A6"] is True
        assert result["S!A7"] is False


def test_evaluator_string_concat() -> None:
    """Test string concatenation operator: &"""
    graph = _make_graph(
        _make_node("S!A1", '="hello"&" world"', None),
        _make_node("S!A2", '="a"&1', None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == "hello world"
        assert result["S!A2"] == "a1"


def test_evaluator_unary_negation() -> None:
    """Test unary minus operator"""
    graph = _make_graph(
        _make_node("S!A1", "=-5", None),
        _make_node("S!A2", "=--5", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == -5.0
        assert result["S!A2"] == 5.0


def test_evaluator_division_by_zero() -> None:
    """Test division by zero returns #DIV/0! error"""
    graph = _make_graph(_make_node("S!A1", "=1/0", None))
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1"])
        assert result["S!A1"] == XlError.DIV


def test_evaluator_operators_with_cell_refs() -> None:
    """Test operators with cell references"""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", None, 3),
        _make_node("S!C1", "=S!A1+S!B1", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!C1"])
        assert result["S!C1"] == 13.0


def test_evaluator_operator_precedence() -> None:
    """Test operator precedence: 1+2*3 = 7, not 9"""
    graph = _make_graph(_make_node("S!A1", "=1+2*3", None))
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1"])
        assert result["S!A1"] == 7.0
