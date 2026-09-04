"""Inverted-tree operators coerce measures instead of raising TypeError (#635)."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _arith_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "operator_coerce.xlsx",
        {
            "Inputs": {
                "B1": "abc",
                "C1": '"',
                "D1": 4.0,
                "B10": 1,
                "C10": 2,
                "D10": 3,
            },
            "Engine": {
                "B1": "=Inputs!B1+1",
                "C1": "=Inputs!C1/100",
                "D1": "=Inputs!D1*2",
                "B10": 1,
                "C10": 2,
                "D10": 3,
            },
            "Outputs": {
                "A1": "=Engine!B1",
                "B1": "=Engine!C1",
                "C1": "=Engine!D1",
                "A10": 1,
                "B10": 2,
                "C10": 3,
            },
        },
    )


def _arith_bindings() -> dict:
    return bindings_document(
        series_entry(
            "inputs",
            "Inputs!B1:D1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "engine_row",
            "Engine!B1:D1",
            layout="series",
            direction="internal",
            header_row=10,
        ),
        series_entry(
            "output_row",
            "Outputs!A1:C1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_emitted_arithmetic_uses_runtime_helpers(tmp_path: Path) -> None:
    modules = generate_inverted(_arith_workbook(tmp_path), _arith_bindings())
    internals = modules["internals.py"]
    assert "xl_add(" in internals
    assert "xl_div(" in internals
    assert "xl_mul(" in internals
    assert "def xl_add" in modules["runtime.py"]
    assert "def xl_div" in modules["runtime.py"]


def _compare_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "operator_compare.xlsx",
        {
            "Engine": {
                "A1": 1,
                "B1": "a",
                "C1": "=A1<B1",
            },
        },
    )


def _compare_bindings() -> dict:
    return bindings_document(
        series_entry("num", "Engine!A1", layout="scalar", direction="input"),
        series_entry("text", "Engine!B1", layout="scalar", direction="input", dtype="string"),
        series_entry("ordered", "Engine!C1", layout="scalar", direction="output"),
    )


def test_emitted_compare_uses_runtime_helper(tmp_path: Path) -> None:
    modules = generate_inverted(_compare_workbook(tmp_path), _compare_bindings())
    assert "xl_lt(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="op_compare")
    assert pkg.compute_ordered(num=1, text="a") == (1.0,)


def test_text_cell_arithmetic_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _arith_workbook(tmp_path)
    cells = ["Engine!B1", "Engine!C1", "Engine!D1"]
    graph = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph).evaluate(cells)
    pkg = load_package(
        generate_inverted(workbook, _arith_bindings()),
        tmp_path,
        name="op_coerce_eval",
    )
    got = pkg.compute_output_row(inputs=("abc", '"', 4.0))
    assert got == (expected["Engine!B1"], expected["Engine!C1"], expected["Engine!D1"])
    assert got == ("#VALUE!", "#VALUE!", 8.0)
