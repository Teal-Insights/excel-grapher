"""Lazy-range migration regressions (#314): exported ranges avoid unused cells.

Exported code models ranges as lazy `Range` values. Selective consumers
(`INDEX`, `MATCH`, lookups) must not evaluate unused cells, while reductions
(`SUM`, `SUMPRODUCT`) preserve Excel's first-error sentinel semantics.
"""

from __future__ import annotations

from typing import Any, cast

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.codegen import CodeGenerator
from tests.integration.utils.parity_harness import (
    assert_codegen_matches_evaluator,
    exec_generated_code,
)


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
    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == 20


def test_index_does_not_evaluate_unused_formula_cells() -> None:
    """Exported INDEX leaves unused sibling formula cells unevaluated."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", "=S!A1+S!A2", None),
        _make_node("S!B1", "=INDEX(S!A1:S!A3, 2)", None),
    )
    generated_results, _code, ns = exec_generated_code(graph, ["S!B1"])
    assert generated_results["S!B1"] == 2

    ns_any = cast("dict[str, Any]", ns)
    merged = dict(ns_any["DEFAULT_INPUTS"])
    ctx = ns_any["EvalContext"](inputs=merged, resolver=ns_any["_resolve_formula"])
    ns_any["xl_cell"](ctx, "S!B1")
    assert "S!A3" not in ctx.cache


def test_match_over_column_slice_of_index() -> None:
    """MATCH consumes an INDEX column view lazily (LIC-DSF pattern)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", None, 2),
        _make_node("S!A2", None, 4),
        _make_node("S!B2", None, 5),
        _make_node("S!C1", "=MATCH(5, INDEX(S!A1:S!B2,,2), 0)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.generated_results["S!C1"] == 2


def test_sum_over_range_with_error_produces_error_code() -> None:
    """SUM over a range containing an error surfaces the first error's code."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", "=1/0", None),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", "=SUM(S!A1:S!A3)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1"])
    assert result.generated_results["S!B1"] == XlError.DIV


def test_and_or_over_lazy_range_codegen_parity() -> None:
    """AND/OR over ranges stay aligned between evaluator and exported code."""
    graph = _make_graph(
        _make_node("S!A1", None, True),
        _make_node("S!A2", None, False),
        _make_node("S!A3", None, True),
        _make_node("S!B1", "=AND(S!A1:S!A3)", None),
        _make_node("S!B2", "=OR(S!A1:S!A3)", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2"])
    assert result.generated_results["S!B1"] is False
    assert result.generated_results["S!B2"] is True


def test_dynamic_offset_range_consumed_by_sum() -> None:
    """Dynamic OFFSET returns a lazy range consumed by SUM."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", None, 0),
        _make_node("S!C1", "=SUM(OFFSET(S!A1, S!B1, 0, 3, 1))", None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.generated_results["S!C1"] == 6.0


def test_sumproduct_criteria_parity_over_ranges() -> None:
    """SUMPRODUCT over aligned ranges with a criteria comparison stays aligned."""
    graph = _make_graph(
        _make_node("S!A1", None, "x"),
        _make_node("S!A2", None, "y"),
        _make_node("S!A3", None, "x"),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!B3", None, 30),
        _make_node("S!C1", '=SUMPRODUCT((S!A1:S!A3="x")*S!B1:S!B3)', None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.generated_results["S!C1"] == 40.0


def test_vlookup_parity_over_range_table() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, "apple"),
        _make_node("S!B1", None, 1),
        _make_node("S!A2", None, "banana"),
        _make_node("S!B2", None, 2),
        _make_node("S!C1", '=VLOOKUP("banana", S!A1:S!B2, 2, FALSE)', None),
    )
    result = assert_codegen_matches_evaluator(graph, ["S!C1"])
    assert result.generated_results["S!C1"] == 2


def test_exported_code_is_numpy_free() -> None:
    """The exported runtime for range consumers embeds no numpy usage."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", "=SUM(S!A1:S!A2)", None),
        _make_node("S!B2", "=INDEX(S!A1:S!A2, 1)", None),
        _make_node("S!B3", "=MATCH(2, S!A1:S!A2, 0)", None),
        _make_node("S!B4", "=SUMPRODUCT(S!A1:S!A2, S!A1:S!A2)", None),
    )
    code = CodeGenerator(graph).generate(["S!B1", "S!B2", "S!B3", "S!B4"])
    assert "import numpy" not in code
    assert "np.array" not in code
    assert "class Range" in code


def test_range_target_boundary_is_materialized() -> None:
    """compute_all returns nested lists (not lazy views) for range targets."""
    graph = _make_graph(
        _make_node("S!A1", None, 10.0),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!C1", "=S!A1*3", None),
    )
    generated_results, code, _ns = exec_generated_code(graph, ["S!B1", "S!C1"])
    assert "'S!B1:C1': xl_range_rows" in code
    value = generated_results["S!B1:C1"]
    assert isinstance(value, list)
    assert value == [[20.0, 30.0]]


def test_lookup_stops_before_trailing_missing_cells() -> None:
    """Exact-match lookups do not force evaluation past the matched row."""
    graph = _make_graph(
        _make_node("S!A1", None, "k1"),
        _make_node("S!B1", None, 100),
        _make_node("S!A2", None, "k2"),
        _make_node("S!B2", None, 200),
        # S!A3/S!B3 intentionally missing from the graph: evaluating them raises.
        _make_node("S!C1", '=VLOOKUP("k1", S!A1:S!B3, 2, FALSE)', None),
    )
    generated_results, _code, _ns = exec_generated_code(graph, ["S!C1"])
    assert generated_results["S!C1"] == 100


def test_missing_cell_still_raises_when_range_is_reduced() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=SUM(S!A1:S!A2)", None),
    )
    code = CodeGenerator(graph).generate(["S!B1"])
    ns: dict[str, Any] = {}
    exec(code, ns)
    with pytest.raises(KeyError):
        ns["compute_all"]()
