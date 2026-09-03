"""Layer A1 — leaf closure is per output (MCVE)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)


def _a1_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a1.xlsx",
        {
            "Inputs": {
                "A1": 60,
                "B1": 3.5,
                "C1": 3.5,
                "B2": 4.0,
                "C2": 4.0,
                "B3": 1,
                "B10": 1,
                "C10": 2,
            },
            "Engine": {
                "B1": "=Inputs!A1",
                "C1": "=B1*(1+Inputs!B2/100)/(1+Inputs!B1/100)",
                "D1": "=C1*(1+Inputs!C2/100)/(1+Inputs!C1/100)",
                "C10": 1,
                "D10": 2,
            },
            "Outputs": {
                "A1": "=Engine!C1",
                "B1": "=Engine!D1",
                "A2": "=Engine!C1",
                "A10": 1,
                "B10": 2,
            },
        },
    )


def _a1_bindings() -> dict:
    return bindings_document(
        series_entry("initial_debt", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "growth",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "interest",
            "Inputs!B2:C2",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry("unused_flag", "Inputs!B3", layout="scalar", direction="input", dtype="int"),
        series_entry("engine_year0", "Engine!B1", layout="scalar", direction="internal"),
        series_entry(
            "engine_path",
            "Engine!C1:D1",
            layout="series",
            direction="internal",
            header_row=10,
        ),
        series_entry(
            "output_path",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
        series_entry("output_year1", "Outputs!A2", layout="scalar", direction="output"),
    )


def test_leaf_closure_excludes_unused_flag_and_ctx(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a1_bindings()), tmp_path, name="a1_pkg")
    required = required_param_names(pkg.compute_output_path)
    names = all_param_names(pkg.compute_output_path)
    assert set(required) == {"initial_debt", "growth", "interest"}
    assert "unused_flag" not in names
    assert "ctx" not in names
    year1_required = required_param_names(pkg.compute_output_year1)
    year1_names = all_param_names(pkg.compute_output_year1)
    assert "unused_flag" not in year1_names
    assert "ctx" not in year1_names
    assert "initial_debt" in year1_required
    assert "growth" in year1_required
    assert "interest" in year1_required


def test_no_setters_or_make_context(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    modules = generate_inverted(workbook, _a1_bindings())
    api = modules["api.py"]
    assert "def make_context" not in api
    assert "def set_" not in api
    pkg = load_package(modules, tmp_path, name="a1_api")
    assert not hasattr(pkg, "make_context")
    assert not hasattr(pkg, "set_growth")
    assert not hasattr(pkg, "set_initial_debt")


def test_numeric_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _a1_workbook(tmp_path)
    bindings = _a1_bindings()
    pkg = load_package(generate_inverted(workbook, bindings), tmp_path, name="a1_num")
    graph = create_dependency_graph(
        workbook,
        ["Outputs!A1", "Outputs!B1"],
        load_values=True,
    )
    evaluator = FormulaEvaluator(graph)
    expected = evaluator.evaluate(["Outputs!A1", "Outputs!B1"])
    got = pkg.compute_output_path(initial_debt=60.0, growth=(3.5, 3.5), interest=(4.0, 4.0))
    assert got == pytest.approx((expected["Outputs!A1"], expected["Outputs!B1"]))
