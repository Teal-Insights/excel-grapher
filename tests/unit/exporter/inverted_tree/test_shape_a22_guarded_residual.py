"""Layer A22 — guard-aware residual legality: may-cycles demote to rung 3 (#616).

Distance-zero residual cycles through at least one guarded edge are may-cycles.
They must not raise at plan time, but instead demote to rung 3 (demand-driven)
where circularity is decided at runtime on the branch actually taken.
"""

from __future__ import annotations

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


def _repro_workbook(tmp_path: Path) -> Path:
    """Exact repro from gh-616: scalar x, y with may-cycle under IF.

    x = IF($A$5=1, y, 10)
    y = x*2
    flag A5 = 0
    """
    return write_workbook(
        tmp_path / "a22_repro.xlsx",
        {
            "Engine": {
                "A5": 0,
                "B2": "=IF($A$5=1, C2, 10)",
                "C2": "=B2*2",
            },
        },
    )


def _repro_bindings() -> dict:
    return bindings_document(
        series_entry("flag", "Engine!A5", layout="scalar", direction="input"),
        series_entry("x", "Engine!B2", layout="scalar", direction="output"),
        series_entry("y", "Engine!C2", layout="scalar", direction="output"),
    )


def _series_may_cycle_workbook(tmp_path: Path) -> Path:
    """Series-shaped variant: x_t = IF(flag, y_t, x_{t-1}+1), y_t = x_t*2."""
    return write_workbook(
        tmp_path / "a22_series.xlsx",
        {
            "Engine": {
                "A1": 0,
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "B2": "=10",
                "C2": "=IF($A$1=1, C3, B2+1)",
                "D2": "=IF($A$1=1, D3, C2+1)",
                "B3": "=B2*2",
                "C3": "=C2*2",
                "D3": "=D2*2",
            },
        },
    )


def _series_may_cycle_bindings() -> dict:
    return bindings_document(
        series_entry("flag", "Engine!A1", layout="scalar", direction="input"),
        series_entry("x", "Engine!B2:D2", layout="series", direction="output", header_row=1),
        series_entry("y", "Engine!B3:D3", layout="series", direction="output", header_row=1),
    )


def _must_cycle_workbook(tmp_path: Path) -> Path:
    """Unconditional must-cycle: x = y*2, y = x+1."""
    return write_workbook(
        tmp_path / "a22_must_cycle.xlsx",
        {
            "Engine": {
                "B2": "=C2*2",
                "C2": "=B2+1",
            },
        },
    )


def _must_cycle_bindings() -> dict:
    return bindings_document(
        series_entry("x", "Engine!B2", layout="scalar", direction="output"),
        series_entry("y", "Engine!C2", layout="scalar", direction="output"),
    )


def test_scalar_may_cycle_demotes_to_rung3_and_evaluates_at_runtime(tmp_path: Path) -> None:
    wb = _repro_workbook(tmp_path)
    bindings = _repro_bindings()

    graph = create_dependency_graph(wb, ["Engine!B2", "Engine!C2"], load_values=True)
    report = graph.cycle_report()
    assert report.has_must_cycles is False
    assert report.has_may_cycles is True

    expected = FormulaEvaluator(graph).evaluate(["Engine!B2", "Engine!C2"])
    assert expected == {"Engine!B2": 10.0, "Engine!C2": 20.0}

    modules = generate_inverted(wb, bindings)
    internals = modules["internals.py"]
    assert "eval_instance" in internals

    pkg = load_package(modules, tmp_path, name="a22_repro_pkg")
    assert pkg.compute_x(flag=0) == (10.0,)
    assert pkg.compute_y(flag=0) == (20.0,)

    # When the guarded branch is actually taken, InstanceCycleError is raised at runtime
    with pytest.raises(pkg.runtime.InstanceCycleError):
        pkg.compute_x(flag=1)


def test_series_may_cycle_demotes_to_rung3_and_evaluates_at_runtime(tmp_path: Path) -> None:
    wb = _series_may_cycle_workbook(tmp_path)
    bindings = _series_may_cycle_bindings()

    cells = ["Engine!B2", "Engine!C2", "Engine!D2", "Engine!B3", "Engine!C3", "Engine!D3"]
    graph = create_dependency_graph(wb, cells, load_values=True)
    report = graph.cycle_report()
    assert report.has_must_cycles is False
    assert report.has_may_cycles is True

    expected = FormulaEvaluator(graph).evaluate(cells)
    assert expected["Engine!B2"] == 10.0
    assert expected["Engine!C2"] == 11.0
    assert expected["Engine!D2"] == 12.0
    assert expected["Engine!B3"] == 20.0
    assert expected["Engine!C3"] == 22.0
    assert expected["Engine!D3"] == 24.0

    modules = generate_inverted(wb, bindings)
    internals = modules["internals.py"]
    assert "eval_instance" in internals

    pkg = load_package(modules, tmp_path, name="a22_series_pkg")
    assert pkg.compute_x(flag=0) == pytest.approx((10.0, 11.0, 12.0))
    assert pkg.compute_y(flag=0) == pytest.approx((20.0, 22.0, 24.0))

    with pytest.raises(pkg.runtime.InstanceCycleError):
        pkg.compute_x(flag=1)


def test_must_cycle_fails_closed_at_plan_time(tmp_path: Path) -> None:
    wb = _must_cycle_workbook(tmp_path)
    bindings = _must_cycle_bindings()

    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        generate_inverted(wb, bindings)
