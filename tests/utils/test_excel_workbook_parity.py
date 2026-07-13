"""Unit tests for cached-workbook parity helpers (no xlwings required)."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import DependencyGraph, Node, XlError, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from tests.utils.excel_workbook_parity import (
    ParityMismatchKind,
    assert_workbook_parity,
    compare_cached_to_evaluator,
    compare_evaluator_to_excel_cache,
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


def test_compare_cached_numeric_match() -> None:
    assert compare_cached_to_evaluator(3.0, 3.0, rtol=1e-5, atol=1e-9) is None


def test_compare_cached_numeric_drift() -> None:
    assert (
        compare_cached_to_evaluator(3.0, 4.0, rtol=1e-5, atol=1e-9)
        == ParityMismatchKind.NUMERIC_DRIFT
    )


def test_compare_cached_error_string_matches_xl_error() -> None:
    assert compare_cached_to_evaluator("#NUM!", XlError.NUM, rtol=1e-5, atol=1e-9) is None


def test_compare_cached_error_string_normalizes_whitespace_and_case() -> None:
    assert compare_cached_to_evaluator("  #num!  ", XlError.NUM, rtol=1e-5, atol=1e-9) is None


def test_compare_cached_xl_error_sentinel_matches() -> None:
    assert compare_cached_to_evaluator(XlError.DIV, XlError.DIV, rtol=1e-5, atol=1e-9) is None


def test_compare_cached_error_code_mismatch() -> None:
    assert (
        compare_cached_to_evaluator("#NUM!", XlError.DIV, rtol=1e-5, atol=1e-9)
        == ParityMismatchKind.XL_ERROR_CODE_MISMATCH
    )


def test_compare_cached_number_vs_evaluator_error() -> None:
    assert (
        compare_cached_to_evaluator(1.0, XlError.NUM, rtol=1e-5, atol=1e-9)
        == ParityMismatchKind.NUMBER_VS_XL_ERROR
    )


def test_compare_cached_error_vs_evaluator_number() -> None:
    assert (
        compare_cached_to_evaluator("#NUM!", 1.0, rtol=1e-5, atol=1e-9)
        == ParityMismatchKind.XL_ERROR_VS_NUMBER
    )


def test_compare_evaluator_to_excel_cache_matching_div_error() -> None:
    graph = _make_graph(_make_node("S!A1", "=1/0", "#DIV/0!"))
    assert compare_evaluator_to_excel_cache(graph, ["S!A1"]) == []


def test_compare_evaluator_to_excel_cache_error_code_mismatch() -> None:
    graph = _make_graph(_make_node("S!A1", "=1/0", "#NUM!"))
    mismatches = compare_evaluator_to_excel_cache(graph, ["S!A1"])
    assert len(mismatches) == 1
    assert mismatches[0].kind == ParityMismatchKind.XL_ERROR_CODE_MISMATCH
    assert mismatches[0].evaluator_result == XlError.DIV


def test_compare_evaluator_to_excel_cache_number_vs_error() -> None:
    graph = _make_graph(_make_node("S!A1", "=1/0", 1.0))
    mismatches = compare_evaluator_to_excel_cache(graph, ["S!A1"])
    assert len(mismatches) == 1
    assert mismatches[0].kind == ParityMismatchKind.NUMBER_VS_XL_ERROR


def test_compare_evaluator_to_excel_cache_error_vs_number() -> None:
    graph = _make_graph(_make_node("S!A1", "=1", "#NUM!"))
    mismatches = compare_evaluator_to_excel_cache(graph, ["S!A1"])
    assert len(mismatches) == 1
    assert mismatches[0].kind == ParityMismatchKind.XL_ERROR_VS_NUMBER
    assert mismatches[0].evaluator_result == 1.0


def test_compare_evaluator_to_excel_cache_leaf_literal_matches_cached_number() -> None:
    graph = _make_graph(_make_node("S!A1", None, 1.0))
    assert compare_evaluator_to_excel_cache(graph, ["S!A1"]) == []


def test_assert_workbook_parity_numeric_unchanged() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 2.0),
        _make_node("S!B1", "=S!A1+1", 3.0),
    )
    graph.add_edge("S!B1", "S!A1")
    assert_workbook_parity(graph, ["S!B1"])


def test_assert_workbook_parity_raises_on_error_code_mismatch() -> None:
    graph = _make_graph(_make_node("S!A1", "=1/0", "#NUM!"))
    with pytest.raises(AssertionError, match="xl_error_code_mismatch"):
        assert_workbook_parity(graph, ["S!A1"])


def test_workbook_parity_error_strings_from_xlsx(tmp_path: Path) -> None:
    wb_path = tmp_path / "errors.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("S")
    ws.write_formula(0, 0, "=1/0", None, "#DIV/0!")
    wb.close()

    graph = create_dependency_graph(wb_path, ["S!A1"], load_values=True)
    node = graph.get_node("S!A1")
    assert node is not None
    assert node.value == "#DIV/0!"

    assert compare_evaluator_to_excel_cache(graph, ["S!A1"]) == []
