"""Finding 2 — mixed formula shapes partition into statements, not fail closed."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _a12_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a12_shapes.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=1",
                "B2": "=A2*2",
                "C2": "=B2+100",
            },
        },
    )


def _a12_vertical_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a12_shapes_vertical.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "A3": 2011,
                "B1": "=1",
                "B2": "=B1*2",
                "B3": "=B2+100",
            },
        },
    )


def _a12_vertical_bindings() -> dict:
    return bindings_document(
        series_entry(
            "path",
            "Engine!B1:B3",
            layout="series",
            direction="output",
            label_column="A",
        ),
    )


def _a12_bindings() -> dict:
    return bindings_document(
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _a12_elementwise_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a12_elementwise.xlsx",
        {
            "Inputs": {"A1": 10.0, "B1": 20.0, "C1": 30.0, "A10": 1, "B10": 2, "C10": 3},
            "Engine": {
                "A1": 1,
                "B1": 2,
                "C1": 3,
                "A2": "=Inputs!A1",
                "B2": "=Inputs!B1*2",
                "C2": "=Inputs!C1+100",
            },
        },
    )


def _a12_elementwise_bindings() -> dict:
    return bindings_document(
        series_entry(
            "values",
            "Inputs!A1:C1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


@pytest.mark.parametrize(
    ("workbook_fn", "bindings_fn", "cells", "pkg_name"),
    [
        (_a12_workbook, _a12_bindings, ["Engine!A2", "Engine!B2", "Engine!C2"], "a12_h"),
        (
            _a12_vertical_workbook,
            _a12_vertical_bindings,
            ["Engine!B1", "Engine!B2", "Engine!B3"],
            "a12_v",
        ),
    ],
    ids=["horizontal", "vertical"],
)
def test_mixed_member_formulas_emit_correct_values(
    tmp_path: Path, workbook_fn, bindings_fn, cells: list[str], pkg_name: str
) -> None:
    workbook = workbook_fn(tmp_path)
    modules = generate_inverted(workbook, bindings_fn())
    assert "formula shape" not in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name=pkg_name)
    assert pkg.compute_path() == pytest.approx((1.0, 2.0, 102.0))

    graph = create_dependency_graph(workbook, cells, load_values=True)
    evaluator = FormulaEvaluator(graph)
    got = evaluator.evaluate(cells)
    assert tuple(got[cell] for cell in cells) == pytest.approx((1.0, 2.0, 102.0))


def test_mixed_elementwise_formulas_emit_correct_values(tmp_path: Path) -> None:
    workbook = _a12_elementwise_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _a12_elementwise_bindings()),
        tmp_path,
        name="a12_elem",
    )
    assert pkg.compute_path(values=(10.0, 20.0, 30.0)) == pytest.approx((10.0, 40.0, 130.0))
