"""INDIRECT lowers from graph edges, not `xl_indirect` (#668)."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.access import classify_producer_access
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _literal_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a27_literal.xlsx",
        {
            "Inputs": {"A1": 42.0},
            "Outputs": {"A1": '=INDIRECT("Inputs!A1")'},
        },
    )


def _literal_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )


def _bound_cell_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a27_bound_cell.xlsx",
        {
            "Inputs": {"A1": 42.0, "B1": "Inputs!A1"},
            "Outputs": {"A1": "=INDIRECT(Inputs!B1)"},
        },
    )


def _bound_cell_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "addr",
            "Inputs!B1",
            layout="scalar",
            direction="input",
            dtype="string",
        ),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )


def _bound_cell_dynamic_refs() -> DynamicRefConfig:
    return DynamicRefConfig.from_constraints({"Inputs!B1": Literal["Inputs!A1"]}, {})


def _series_member_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a27_series.xlsx",
        {
            "Inputs": {
                "B1": 10.0,
                "C1": 20.0,
                "D1": 30.0,
                "B10": 1,
                "C10": 2,
                "D10": 3,
            },
            "Outputs": {"A1": '=INDIRECT("Inputs!C1")'},
        },
    )


def _series_member_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "src",
            "Inputs!B1:D1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )


def _assert_no_xl_indirect(modules: dict[str, str]) -> None:
    internals = modules["internals.py"]
    runtime = modules["runtime.py"]
    assert "xl_indirect" not in internals
    assert "def xl_indirect" not in runtime


def test_literal_address_matches_evaluator(tmp_path: Path) -> None:
    workbook = _literal_workbook(tmp_path)
    document = _literal_bindings()
    modules = generate_inverted(workbook, document)
    _assert_no_xl_indirect(modules)
    pkg = load_package(modules, tmp_path, name="a27_lit")
    _catalog, _deps, graph = inverted_graph_parts(workbook, document)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])["Outputs!A1"]
    assert pkg.compute_out(src=42.0) == (pytest.approx(expected),)
    assert "src" in all_param_names(pkg.compute_out)


def test_bound_cell_address_matches_evaluator(tmp_path: Path) -> None:
    workbook = _bound_cell_workbook(tmp_path)
    document = _bound_cell_bindings()
    refs = _bound_cell_dynamic_refs()
    modules = generate_inverted(workbook, document, dynamic_refs=refs)
    _assert_no_xl_indirect(modules)
    pkg = load_package(modules, tmp_path, name="a27_cell")
    _catalog, _deps, graph = inverted_graph_parts(workbook, document, dynamic_refs=refs)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])["Outputs!A1"]
    assert pkg.compute_out(src=42.0) == (pytest.approx(expected),)


def test_literal_into_series_emits_xl_at(tmp_path: Path) -> None:
    workbook = _series_member_workbook(tmp_path)
    document = _series_member_bindings()
    modules = generate_inverted(workbook, document)
    _assert_no_xl_indirect(modules)
    assert "xl_at(" in modules["internals.py"]
    assert "xl_at(src, 1)" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a27_series")
    _catalog, _deps, graph = inverted_graph_parts(workbook, document)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])["Outputs!A1"]
    assert pkg.compute_out(src=(10.0, 20.0, 30.0)) == (pytest.approx(expected),)


def test_literal_access_is_static(tmp_path: Path) -> None:
    workbook = _literal_workbook(tmp_path)
    catalog, _deps, graph = inverted_graph_parts(workbook, _literal_bindings())
    access = classify_producer_access(catalog.get("out"), catalog.get("src"), catalog, graph)
    assert access.row.kind == "whole"
    assert access.col.kind == "whole"


def test_unbound_indirect_target_fails_closed_naming_host(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a27_unbound.xlsx",
        {
            "Inputs": {"A1": 1.0, "Z99": 9.0},
            "Outputs": {"A1": '=INDIRECT("Inputs!Z99")'},
        },
    )
    document = bindings_document(
        series_entry("src", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match="Outputs!A1") as exc:
        generate_inverted(workbook, document)
    assert "INDIRECT" in str(exc.value) or "bound series" in str(exc.value)
