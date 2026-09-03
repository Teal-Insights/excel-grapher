"""Rung 3 is the floor: backward recursion must emit, not fail closed."""

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


def _horizontal_terminal_workbook(tmp_path: Path) -> Path:
    """`value_t = value_{t+1} * 0.9` with a terminal seed (common TV models)."""
    return write_workbook(
        tmp_path / "a19_terminal.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=B2*0.9",
                "B2": "=C2*0.9",
                "C2": "=100",
            },
        },
    )


def _horizontal_terminal_bindings() -> dict:
    return bindings_document(
        series_entry(
            "value",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _vertical_terminal_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a19_terminal_vertical.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "A3": 2011,
                "B1": "=B2*0.9",
                "B2": "=B3*0.9",
                "B3": "=100",
            },
        },
    )


def _vertical_terminal_bindings() -> dict:
    return bindings_document(
        series_entry(
            "value",
            "Engine!B1:B3",
            layout="series",
            direction="output",
            label_column="A",
        ),
    )


@pytest.mark.parametrize(
    ("workbook_fn", "bindings_fn", "cells", "pkg_name"),
    [
        (
            _horizontal_terminal_workbook,
            _horizontal_terminal_bindings,
            ["Engine!A2", "Engine!B2", "Engine!C2"],
            "a19_h",
        ),
        (
            _vertical_terminal_workbook,
            _vertical_terminal_bindings,
            ["Engine!B1", "Engine!B2", "Engine!B3"],
            "a19_v",
        ),
    ],
    ids=["horizontal", "vertical"],
)
def test_backward_recursion_emits_rung3_and_matches_evaluator(
    tmp_path: Path,
    workbook_fn,
    bindings_fn,
    cells: list[str],
    pkg_name: str,
) -> None:
    workbook = workbook_fn(tmp_path)
    modules = generate_inverted(workbook, bindings_fn())
    internals = modules["internals.py"]
    assert "eval_instance" in internals
    assert "non-lag cell" not in internals
    pkg = load_package(modules, tmp_path, name=pkg_name)
    graph = create_dependency_graph(workbook, cells, load_values=True)
    assert graph.cycle_report().has_must_cycles is False
    expected = FormulaEvaluator(graph).evaluate(cells)
    got = pkg.compute_value()
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((81.0, 90.0, 100.0))
