"""OFFSET: evaluator ↔ export parity on synthetic graphs, plus known gaps.

Three-arg cell OFFSET goes through `evaluate_targets` (evaluator and inverted-tree
export). Height/width is fail-closed at emit (`InvertedTreeExportError`). OFFSET
past the sheet is `#REF!` in Excel/evaluator and `#VALUE!` in the bound model.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.grapher.writeback import write_workbook
from excel_grapher.series_bindings import validate_bindings_document
from tests.integration.utils.parity_harness import (
    _bindings_for_graph,
    _write_axis_keys,
    evaluate_targets,
)
from tests.unit.exporter.inverted_tree.helpers import invoke_public_compute, load_package


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


def test_offset_parity_static_single_cell() -> None:
    """OFFSET with constant offsets resolves to single cell; evaluator and generated code agree."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!B1", "=OFFSET(S!A1, 1, 0)", None),
    )
    results = evaluate_targets(graph, ["S!B1"])
    assert results["S!B1"] == 20


def test_offset_parity_static_range_sum() -> None:
    """OFFSET with constant height/width consumed by SUM (evaluator; emit has no size args)."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!C1", "=SUM(OFFSET(S!A1, 0, 0, 2, 2))", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!C1"])["S!C1"] == 33.0


def test_offset_parity_dynamic_row_offset() -> None:
    """OFFSET with row offset from cell (dynamic); evaluator and generated code agree."""
    graph = _make_graph(
        _make_node("S!A1", None, 5),
        _make_node("S!A2", None, 10),
        _make_node("S!A3", None, 15),
        _make_node("S!B1", None, 1),
        _make_node("S!C1", "=OFFSET(S!A1, S!B1, 0)", None),
    )
    results = evaluate_targets(graph, ["S!C1"])
    assert results["S!C1"] == 10


def test_offset_parity_negative_offset() -> None:
    """OFFSET with negative row offset; both runtimes return same value."""
    graph = _make_graph(
        _make_node("S!A1", None, 100),
        _make_node("S!A2", None, 200),
        _make_node("S!A3", None, 300),
        _make_node("S!B1", "=OFFSET(S!A3, -2, 0)", None),
    )
    results = evaluate_targets(graph, ["S!B1"])
    assert results["S!B1"] == 100


def test_offset_parity_invalid_returns_ref_error() -> None:
    """OFFSET past the worksheet returns `#REF!` in the evaluator."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", "=OFFSET(S!A1, -1, 0)", None),
    )
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"])["S!B1"] == XlError.REF


def test_offset_past_sheet_bound_export_raises_value(tmp_path: Path) -> None:
    """Bound inverted-tree OFFSET past the series is `#VALUE!`, not Excel `#REF!`."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", "=OFFSET(S!A1, -1, 0)", None),
    )
    graph.sheet_order = ["S"]
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!B1"])["S!B1"] == XlError.REF
    workbook = tmp_path / "offset_ref.xlsx"
    write_workbook(graph, workbook)
    document, axis_plan = _bindings_for_graph(graph, ["S!B1"])
    _write_axis_keys(workbook, axis_plan)
    modules = generate_inverted_tree_modules(
        graph,
        series_bindings=validate_bindings_document(document),
        bindings_workbook=workbook,
    )
    pkg = load_package(modules, tmp_path, name="offset_ref_div")
    outputs = [series for series in document["series"] if "output" in series]
    assert len(outputs) == 1
    # Series-member emit stores the code rather than aborting the compute.
    compute = getattr(pkg, f"compute_{outputs[0]['id']}")
    assert invoke_public_compute(pkg, compute, {}) == "#VALUE!"


def test_offset_height_width_evaluate_targets_is_fail_closed() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!C1", "=SUM(OFFSET(S!A1, 0, 0, 2, 2))", None),
    )
    with pytest.raises(InvertedTreeExportError, match="height/width"):
        evaluate_targets(graph, ["S!C1"])
