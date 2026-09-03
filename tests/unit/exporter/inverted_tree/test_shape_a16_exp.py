"""Layer A16 — Excel `EXP` must import from inverted-tree `runtime.py` (#606)."""

from __future__ import annotations

import math
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _exp_workbook(tmp_path: Path, formula: str, *, x: float = 0) -> Path:
    return write_workbook(
        tmp_path / "a16_exp.xlsx",
        {"Engine": {"A1": x, "B1": formula}},
    )


def _exp_bindings() -> dict:
    return bindings_document(
        series_entry("x", "Engine!A1", layout="scalar", direction="input"),
        series_entry("exp_x", "Engine!B1", layout="scalar", direction="output"),
    )


def test_exp_emits_runtime_helper_and_imports(tmp_path: Path) -> None:
    workbook = _exp_workbook(tmp_path, "=EXP(A1)")
    modules = generate_inverted(workbook, _exp_bindings())
    assert "xl_exp(" in modules["internals.py"]
    assert "def xl_exp" in modules["runtime.py"]
    pkg = load_package(modules, tmp_path, name="a16_exp")
    assert pkg.compute_exp_x(x=0) == pytest.approx((1.0,))
    assert pkg.compute_exp_x(x=1) == pytest.approx((math.e,))


def test_exp_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _exp_workbook(tmp_path, "=EXP(A1)", x=1)
    pkg = load_package(generate_inverted(workbook, _exp_bindings()), tmp_path, name="a16_exp_eval")
    graph = create_dependency_graph(workbook, ["Engine!B1"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Engine!B1"])
    assert pkg.compute_exp_x(x=1) == pytest.approx((expected["Engine!B1"],))


def test_exp_overflow_is_num_error_measure(tmp_path: Path) -> None:
    workbook = _exp_workbook(tmp_path, "=EXP(A1)", x=1000)
    pkg = load_package(generate_inverted(workbook, _exp_bindings()), tmp_path, name="a16_exp_ovf")
    assert pkg.compute_exp_x(x=1000) == ("#NUM!",)


def test_unknown_excel_function_fails_closed(tmp_path: Path) -> None:
    workbook = _exp_workbook(tmp_path, "=LN(A1)")
    with pytest.raises(InvertedTreeExportError, match="LN"):
        generate_inverted(workbook, _exp_bindings())
