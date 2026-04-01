"""Tests for declared structural blank ranges (issue #39)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node, create_dependency_graph
from excel_grapher.evaluator.name_utils import parse_address
from excel_grapher.grapher.blank_ranges import (
    cell_in_blank_ranges,
    normalize_blank_range_specs,
    parse_blank_range_spec,
)
from tests.evaluator.parity_harness import assert_codegen_matches_evaluator


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


def test_parse_blank_range_spec_single_cell() -> None:
    assert parse_blank_range_spec("Sheet1!B2") == ("Sheet1", 2, 2, 2, 2)


def test_parse_blank_range_spec_rectangle() -> None:
    assert parse_blank_range_spec("Sheet1!B2:D4") == ("Sheet1", 2, 2, 4, 4)


def test_parse_blank_range_spec_quoted_sheet() -> None:
    assert parse_blank_range_spec("'My Sheet'!A1:C2") == ("My Sheet", 1, 1, 2, 3)


def test_normalize_blank_range_specs_rejects_str() -> None:
    with pytest.raises(TypeError):
        normalize_blank_range_specs("Sheet1!A1")  # type: ignore[arg-type]


def test_cell_in_blank_ranges() -> None:
    rects = normalize_blank_range_specs(["S!A2:B3"])
    assert cell_in_blank_ranges("S", 2, 1, rects)
    assert cell_in_blank_ranges("S", 3, 2, rects)
    assert not cell_in_blank_ranges("S", 1, 1, rects)
    assert not cell_in_blank_ranges("T", 2, 1, rects)


def test_create_dependency_graph_skips_blank_range_nodes(tmp_path: Path) -> None:
    path = tmp_path / "blank_range.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["B1"].value = 20
    ws["D1"].value = "=INDEX(A1:B3,1,1)"
    ws["E1"].value = "=INDEX(A1:B3,3,2)"
    wb.save(path)
    wb.close()

    blank = ("Sheet1!A2:B3",)
    graph = create_dependency_graph(
        path, ["Sheet1!D1", "Sheet1!E1"], load_values=True, blank_ranges=blank
    )

    for addr in ("Sheet1!A2", "Sheet1!B2", "Sheet1!A3", "Sheet1!B3"):
        assert addr not in graph

    assert "Sheet1!A1" in graph
    assert "Sheet1!B1" in graph
    assert "Sheet1!D1" in graph
    assert "Sheet1!E1" in graph

    with FormulaEvaluator(graph, blank_ranges=blank) as ev:
        assert ev._evaluate_cell("Sheet1!D1") == 10  # noqa: SLF001
        assert ev._evaluate_cell("Sheet1!E1") == 0  # noqa: SLF001


def test_missing_cell_outside_declared_blank_still_keyerror() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", "=S!Z99", None))

    with FormulaEvaluator(graph) as ev, pytest.raises(KeyError):
        ev._evaluate_cell("S!A1")  # noqa: SLF001

    with FormulaEvaluator(graph, blank_ranges=("S!Z99",)) as ev2:
        assert ev2._evaluate_cell("S!A1") == 0  # noqa: SLF001


def test_blank_range_codegen_compact_and_parity(tmp_path: Path) -> None:
    path = tmp_path / "blank_range_parity.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 10
    ws["B1"].value = 20
    ws["D1"].value = "=INDEX(A1:B3,1,1)"
    ws["E1"].value = "=INDEX(A1:B3,3,2)"
    wb.save(path)
    wb.close()

    blank = ("Sheet1!A2:B3",)
    graph = create_dependency_graph(
        path, ["Sheet1!D1", "Sheet1!E1"], load_values=True, blank_ranges=blank
    )

    from excel_grapher.evaluator.codegen import CodeGenerator

    code = CodeGenerator(graph).generate(["Sheet1!D1", "Sheet1!E1"], blank_ranges=blank)
    assert "_BLANK_RANGE_RECTS" in code
    assert "def cell_sheet1_a2(" not in code
    assert "_blank_structural_cell" in code

    assert_codegen_matches_evaluator(graph, ["Sheet1!D1", "Sheet1!E1"], blank_ranges=blank)
