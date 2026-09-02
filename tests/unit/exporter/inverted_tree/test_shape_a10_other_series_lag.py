"""Layer A10 — a series cell may read another series at t and t-1 (#601)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _lag_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a10_lag.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+1",
                "C2": "=B2+1",
                "B3": "=IF(B2>A2,1,IF(B2<A2,2,))",
                "C3": "=IF(C2>B2,1,IF(C2<B2,2,))",
            },
        },
    )


def _lag_bindings() -> dict:
    return bindings_document(
        series_entry(
            "debt",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        series_entry(
            "direction",
            "Engine!B3:C3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def _non_lag_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a10_nonlag.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": 100.0,
                "B2": 101.0,
                "C2": 102.0,
                "B3": "=A2+C2",
                "C3": "=B2+C2",
            },
        },
    )


def _non_lag_bindings() -> dict:
    return bindings_document(
        series_entry(
            "debt",
            "Engine!A2:C2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "direction",
            "Engine!B3:C3",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def test_other_series_lag_emits_without_two_position_error(tmp_path: Path) -> None:
    workbook = _lag_workbook(tmp_path)
    modules = generate_inverted(workbook, _lag_bindings())
    internals = modules["internals.py"]
    assert "two positions" not in internals
    direction_fn = internals[internals.index("def direction") :]
    assert "debt[i]" in direction_fn
    assert "debt[i + 1]" in direction_fn
    assert "prior:" not in direction_fn
    pkg = load_package(modules, tmp_path, name="a10_lag")
    assert "debt" in all_param_names(pkg.internals.direction)
    assert pkg.internals.direction((100.0, 101.0, 102.0)) == (1.0, 1.0)
    assert pkg.internals.direction((102.0, 101.0, 100.0)) == (2.0, 2.0)
    assert pkg.internals.direction((100.0, 102.0, 101.0)) == (1.0, 2.0)


def test_other_series_lag_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _lag_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _lag_bindings()), tmp_path, name="a10_num")
    graph = create_dependency_graph(
        workbook,
        ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"],
        load_values=True,
    )
    expected = FormulaEvaluator(graph).evaluate(
        ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"]
    )
    debt = (expected["Engine!A2"], expected["Engine!B2"], expected["Engine!C2"])
    got = pkg.internals.direction(debt)
    assert got == pytest.approx((expected["Engine!B3"], expected["Engine!C3"]))


def test_non_adjacent_two_positions_still_fail_closed(tmp_path: Path) -> None:
    workbook = _non_lag_workbook(tmp_path)
    with pytest.raises(InvertedTreeExportError, match="two positions"):
        generate_inverted(workbook, _non_lag_bindings())
