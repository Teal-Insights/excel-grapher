"""Layer A11 — period-lag zippers are a cell DAG and a series SCC (#603)."""

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


def _zipper_workbook(tmp_path: Path) -> Path:
    """`adj_t` reads `debt_{t-1}`; `debt_t` reads `adj_t` and `debt_{t-1}`."""
    return write_workbook(
        tmp_path / "a11_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B3": "=A2",
                "B2": "=A2+B3",
                "C3": "=B2",
                "C2": "=B2+C3",
            },
        },
    )


def _zipper_bindings() -> dict:
    return bindings_document(
        series_entry(
            "debt",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        series_entry(
            "adjustment",
            "Engine!B3:C3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def _simultaneous_workbook(tmp_path: Path) -> Path:
    """Same-year `debt_t ↔ adj_t` (Excel circular ref), not a lag zipper."""
    return write_workbook(
        tmp_path / "a11_simultaneous.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+B3",
                "C2": "=B2+C3",
                "B3": "=B2-A2",
                "C3": "=C2-B2",
            },
        },
    )


def test_lag_zipper_emits_joint_year_loop(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    modules = generate_inverted(workbook, _zipper_bindings())
    internals = modules["internals.py"]
    api = modules["api.py"]
    assert "cyclic formula-series" not in internals
    assert "for i in range(" in internals
    assert internals.count("for i in range(") == 1
    pkg = load_package(modules, tmp_path, name="a11_zip")
    got = pkg.compute_debt()
    assert got == pytest.approx((100.0, 200.0, 400.0))
    assert "scan_debt_adjustment" in internals
    assert "scan_debt_adjustment" in api
    assert "internals.scan_debt_adjustment" in api


def test_lag_zipper_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _zipper_bindings()), tmp_path, name="a11_num")
    graph = create_dependency_graph(
        workbook,
        ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"],
        load_values=True,
    )
    expected = FormulaEvaluator(graph).evaluate(
        ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"]
    )
    got = pkg.compute_debt()
    assert got == pytest.approx(
        (expected["Engine!A2"], expected["Engine!B2"], expected["Engine!C2"])
    )
    adj = pkg.internals.adjustment()
    assert adj == pytest.approx((expected["Engine!B3"], expected["Engine!C3"]))


def test_same_year_cell_cycle_still_fail_closed(tmp_path: Path) -> None:
    workbook = _simultaneous_workbook(tmp_path)
    with pytest.raises(InvertedTreeExportError, match="cell cycle"):
        generate_inverted(workbook, _zipper_bindings())
