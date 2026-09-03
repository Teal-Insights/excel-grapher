"""Layer A18 — a spliced series copies the last prefix source, then another series (#608)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _time_series(
    series_id: str,
    data_range: str,
    *,
    output: bool = False,
    internal: bool = False,
) -> dict:
    if output:
        direction = "output"
    elif internal:
        direction = "internal"
    else:
        direction = "input"
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction=direction,
        header_row=1,
    )


def _splice_workbook(tmp_path: Path) -> Path:
    """Path `D:G` is last growth year, then a 3-year trajectory."""
    return write_workbook(
        tmp_path / "a18_splice.xlsx",
        {
            "Engine": {
                "B1": 2009,
                "C1": 2010,
                "D1": 2011,
                "E1": 2012,
                "F1": 2013,
                "G1": 2014,
                "B2": 100,
                "C2": 110,
                "D2": 121,
                "C3": "=C2/B2",
                "D3": "=D2/C2",
                "E5": "=0.02",
                "F5": "=0.03",
                "G5": "=0.04",
                "D4": "=D3",
                "E4": "=E5",
                "F4": "=F5",
                "G4": "=G5",
                "C7": "=C3",
                "D6": "=D4",
                "E6": "=E4",
                "F6": "=F4",
                "G6": "=G4",
            },
        },
    )


def _splice_bindings() -> dict:
    return bindings_document(
        _time_series("gdp", "Engine!B2:D2"),
        _time_series("growth", "Engine!C3:D3", internal=True),
        _time_series("trajectory", "Engine!E5:G5", internal=True),
        _time_series("path", "Engine!D4:G4", internal=True),
        _time_series("result", "Engine!D6:G6", output=True),
        series_entry(
            "growth_keep",
            "Engine!C7",
            layout="scalar",
            direction="output",
        ),
    )


def test_splice_partitions_path_into_prefix_and_tail(tmp_path: Path) -> None:
    catalog, deps, _graph = inverted_graph_parts(_splice_workbook(tmp_path), _splice_bindings())
    path = catalog.get("path")
    assert [(stmt.start, stmt.stop) for stmt in path.statements] == [(0, 1), (1, 4)]
    assert "growth" not in deps["path"].aligned_ids
    assert "trajectory" not in deps["path"].aligned_ids


def test_splice_indexes_last_growth_then_trajectory(tmp_path: Path) -> None:
    workbook = _splice_workbook(tmp_path)
    modules = generate_inverted(workbook, _splice_bindings())
    internals = modules["internals.py"]
    assert "as_measure(growth[i + 1])" in internals
    assert "as_measure(trajectory[i - 1])" in internals
    assert "if i < 1:" in internals
    pkg = load_package(modules, tmp_path, name="a18_splice")
    got = pkg.compute_result(gdp=(100.0, 110.0, 121.0))
    assert got == pytest.approx((121 / 110, 0.02, 0.03, 0.04))
    assert pkg.compute_growth_keep(gdp=(100.0, 110.0, 121.0)) == pytest.approx((110 / 100,))


def test_splice_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _splice_workbook(tmp_path)
    targets = ["Engine!D6", "Engine!E6", "Engine!F6", "Engine!G6"]
    pkg = load_package(generate_inverted(workbook, _splice_bindings()), tmp_path, name="a18_eval")
    graph = create_dependency_graph(workbook, targets, load_values=True)
    expected = FormulaEvaluator(graph).evaluate(targets)
    got = pkg.compute_result(gdp=(100.0, 110.0, 121.0))
    assert got == pytest.approx(tuple(expected[addr] for addr in targets))
