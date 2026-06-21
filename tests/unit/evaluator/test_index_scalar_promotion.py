"""INDEX single-cell results must scalar-promote in value contexts (issue #264)."""

from __future__ import annotations

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator

_K16_LOOKUP_FORMULA = '=IFERROR(NUMBERVALUE(TEXT(INDEX(PL!E5:PL!E6,MATCH(PL!K5,PL!A5:PL!A6,0)),"0.00"),".",","),"N/A")'
_TEXT_LOOKUP_FORMULA = '=TEXT(INDEX(PL!E5:PL!E6,MATCH(PL!K5,PL!A5:PL!A6,0)),"0.00")'


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


def _make_lookup_graph(*, target: str, formula: str) -> DependencyGraph:
    graph = DependencyGraph()
    for node in (
        _make_node("PL!K5", None, "PRD-001"),
        _make_node("PL!A5", None, "PRD-001"),
        _make_node("PL!A6", None, "PRD-002"),
        _make_node("PL!E5", None, 1499.0),
        _make_node("PL!E6", None, 999.0),
        _make_node(target, formula, None),
    ):
        graph.add_node(node)
    return graph


def test_text_index_match_promotes_single_cell_index_to_scalar() -> None:
    """``TEXT(INDEX(..., MATCH(...)), ...)`` formats a scalar, not a stringified array."""
    graph = _make_lookup_graph(target="PL!K16", formula=_TEXT_LOOKUP_FORMULA)
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("PL!K16") == "1499.00"


def test_numbervalue_text_index_match_returns_numeric_price() -> None:
    """``NUMBERVALUE(TEXT(INDEX(...)))`` returns the looked-up price (K16 shape)."""
    graph = _make_lookup_graph(target="PL!K16", formula=_K16_LOOKUP_FORMULA)
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("PL!K16") == 1499.0


def test_numbervalue_text_index_match_eval_codegen_parity() -> None:
    """Evaluator and export agree on the K16 ``NUMBERVALUE(TEXT(INDEX(...)))`` chain."""
    graph = _make_lookup_graph(target="PL!K16", formula=_K16_LOOKUP_FORMULA)
    result = assert_codegen_matches_evaluator(graph, ["PL!K16"])
    assert result.evaluator_results["PL!K16"] == 1499.0
    assert result.generated_results["PL!K16"] == 1499.0
