"""Unit tests for outermost embedded-conditional span selection."""

from __future__ import annotations

from excel_grapher.grapher.builder import _outermost_embedded_conditional_spans


def test_arithmetic_embedding_selects_the_IF_span() -> None:
    formula = "=1+IF(B1=1,C1,D1)"
    spans = _outermost_embedded_conditional_spans(formula)
    assert len(spans) == 1
    start, end = spans[0]
    assert formula[start:end] == "IF(B1=1,C1,D1)"


def test_nested_IF_inside_arithmetic_selects_only_outermost() -> None:
    formula = "=1+IF(A1=1,IF(B1=1,C1,D1),0)"
    spans = _outermost_embedded_conditional_spans(formula)
    assert len(spans) == 1
    start, end = spans[0]
    assert formula[start:end] == "IF(A1=1,IF(B1=1,C1,D1),0)"


def test_sibling_embedded_IFs_are_both_selected() -> None:
    formula = "=IF(B1=1,C1,D1)+IF(E1=1,F1,G1)"
    spans = _outermost_embedded_conditional_spans(formula)
    texts = sorted(formula[a:b] for a, b in spans)
    assert texts == ["IF(B1=1,C1,D1)", "IF(E1=1,F1,G1)"]


def test_top_level_conditional_is_not_treated_as_embedded() -> None:
    assert _outermost_embedded_conditional_spans("=IF(B1=1,C1,D1)") == []
    assert _outermost_embedded_conditional_spans("=IFS(B1=1,C1,TRUE,0)") == []


def test_conditional_inside_OFFSET_is_skipped() -> None:
    formula = "=A1+OFFSET(IF(B1=1,C1,D1),0,0)"
    assert _outermost_embedded_conditional_spans(formula) == []


def test_IF_sibling_to_OFFSET_is_still_selected() -> None:
    formula = "=IF(B1=1,C1,D1)+OFFSET(A1,0,0)"
    spans = _outermost_embedded_conditional_spans(formula)
    assert len(spans) == 1
    start, end = spans[0]
    assert formula[start:end] == "IF(B1=1,C1,D1)"
