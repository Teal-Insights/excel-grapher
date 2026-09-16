from __future__ import annotations

import sys
from typing import Any

import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from tests.integration.utils.parity_harness import _bindings_for_graph, evaluate_targets
from tests.unit.exporter.inverted_tree.helpers import load_package


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


def _constant_ranges(document: dict[str, Any]) -> set[str]:
    return {str(series["data_range"]) for series in document["series"] if "constant" in series}


def _formula_series(document: dict[str, Any]) -> list[dict[str, Any]]:
    return [series for series in document["series"] if "output" in series or "internal" in series]


def test_evaluator_simple_arithmetic() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1+S!A2", None),
    )
    results = evaluate_targets(graph, ["S!B1"])
    assert results["S!B1"] == 15.0


def test_evaluator_row_with_reference_and_offset() -> None:
    graph = _make_graph(
        _make_node("S!B3", None, 10),
        _make_node("S!C1", None, 2),
        _make_node("S!C2", None, 0),
        _make_node("S!A1", "=ROW(S!B3)", None),
        _make_node("S!A2", "=ROW(OFFSET(S!B3, S!C1, S!C2))", None),
    )
    results = evaluate_targets(graph, ["S!A1"])
    assert results["S!A1"] == 3
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate(["S!A2"])["S!A2"] == 5


def test_evaluator_row_without_reference() -> None:
    graph = _make_graph(_make_node("S!E12", "=ROW()", None))
    results = evaluate_targets(graph, ["S!E12"])
    assert results["S!E12"] == 12


def test_evaluator_column_and_columns_with_offset() -> None:
    graph = _make_graph(
        _make_node("S!D4", None, 10),
        _make_node("S!C1", None, 0),
        _make_node("S!C2", None, 2),
        _make_node("S!A1", "=COLUMN(S!D4)", None),
        _make_node("S!A2", "=COLUMNS(S!D4)", None),
        _make_node("S!A3", "=COLUMN(OFFSET(S!D4, S!C1, S!C2))", None),
        _make_node("S!A4", "=COLUMNS(OFFSET(S!D4, S!C1, S!C2, 1, 3))", None),
    )
    results = evaluate_targets(graph, ["S!A1", "S!A2"])
    assert results["S!A1"] == 4
    assert results["S!A2"] == 1
    with FormulaEvaluator(graph) as ev:
        rest = ev.evaluate(["S!A3", "S!A4"])
    assert rest["S!A3"] == 6
    assert rest["S!A4"] == 3


def test_evaluate_targets_quoted_sheet_names() -> None:
    graph = _make_graph(
        _make_node("Product Lookup!A1", None, 10),
        _make_node("Product Lookup!B1", "='Product Lookup'!A1+5", None),
    )
    results = evaluate_targets(graph, ["Product Lookup!B1"])
    assert results["Product Lookup!B1"] == 15.0


def test_evaluate_targets_maps_export_errors_to_evaluator_sentinels() -> None:
    graph = _make_graph(_make_node("S!A1", "=1/0", None))
    results = evaluate_targets(graph, ["S!A1"])
    assert results["S!A1"] == XlError.DIV


def test_evaluate_targets_keeps_error_looking_strings() -> None:
    graph = _make_graph(_make_node("S!A1", '="#VALUE!"', None))
    results = evaluate_targets(graph, ["S!A1"])
    assert results["S!A1"] == "#VALUE!"


def test_evaluate_targets_does_not_swallow_export_errors() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", None, 10),
        _make_node("S!B2", None, 20),
        _make_node("S!C1", "=SUM(OFFSET(S!A1, 0, 0, 2, 2))", None),
    )
    with pytest.raises(InvertedTreeExportError, match="height/width"):
        evaluate_targets(graph, ["S!C1"])


def test_evaluate_targets_does_not_mutate_sheet_order() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1+S!A2", None),
    )
    assert graph.sheet_order is None
    evaluate_targets(graph, ["S!B1"])
    assert graph.sheet_order is None


def test_evaluate_targets_unloads_parity_modules() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1+S!A2", None),
    )
    before = {key for key in sys.modules if key.startswith("parity_")}
    evaluate_targets(graph, ["S!B1"])
    after = {key for key in sys.modules if key.startswith("parity_")}
    assert after == before


def test_load_package_pops_sys_path(tmp_path) -> None:
    modules = {
        "__init__.py": "",
        "api.py": "",
        "internals.py": "",
        "runtime.py": "",
        "data.py": "",
        "validation.py": "",
    }
    path_str = str(tmp_path)
    assert path_str not in sys.path
    load_package(modules, tmp_path, name="parity_path_probe")
    assert path_str not in sys.path


def test_bindings_omit_formulas_outside_the_demand_cone() -> None:
    graph = _make_graph(
        _make_node("S!B3", None, 10),
        _make_node("S!C1", None, 2),
        _make_node("S!C2", None, 0),
        _make_node("S!A1", "=ROW(S!B3)", None),
        _make_node("S!A2", "=ROW(OFFSET(S!B3, S!C1, S!C2))", None),
    )
    document, _axis = _bindings_for_graph(graph, ["S!A1"])
    formula_ranges = {str(series["data_range"]) for series in _formula_series(document)}
    assert "S!A1" in formula_ranges
    assert "S!A2" not in formula_ranges


def test_bindings_dense_rect_is_one_covering_series() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", None, 2),
        _make_node("S!A2", None, 3),
        _make_node("S!B2", None, 4),
        _make_node("S!C1", "=INDEX(S!A1:B2, 2, 2)", None),
    )
    document, _axis = _bindings_for_graph(graph, ["S!C1"])
    constants = [series for series in document["series"] if "constant" in series]
    assert len(constants) == 1
    assert constants[0]["layout"] == "matrix"
    assert constants[0]["data_range"] == "S!A1:B2"


def test_bindings_sparse_bbox_falls_back_to_column_runs() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A3", None, 3),
        _make_node("S!C1", None, 9),
        _make_node("S!D1", "=S!A1+S!A3+S!C1", None),
    )
    document, _axis = _bindings_for_graph(graph, ["S!D1"])
    assert _constant_ranges(document) == {"S!A1", "S!A3", "S!C1"}


def test_bindings_formula_in_the_middle_splits_leaf_column() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", "=S!A1", None),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", "=SUM(S!A1:A3)", None),
    )
    document, _axis = _bindings_for_graph(graph, ["S!B1"])
    assert _constant_ranges(document) == {"S!A1", "S!A3"}
    formula_ranges = {str(series["data_range"]) for series in _formula_series(document)}
    assert formula_ranges == {"S!B1", "S!A2"}


def test_bindings_mixed_leaf_types_use_float_dtype() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!B1", None, "x"),
        _make_node("S!A2", None, 2),
        _make_node("S!B2", None, 3),
        _make_node("S!C1", "=INDEX(S!A1:B2, 1, 1)", None),
    )
    document, _axis = _bindings_for_graph(graph, ["S!C1"])
    constants = [series for series in document["series"] if "constant" in series]
    assert len(constants) == 1
    assert constants[0]["structure"]["measure"]["dtype"] == "float"
