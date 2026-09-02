"""Layer A13 — residual identity can flip across years without a cell cycle."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import plan_fused_scc
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _time_series(series_id: str, data_range: str, *, output: bool = True) -> dict:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction="output" if output else "internal",
        header_row=1,
    )


def _two_series_workbook(tmp_path: Path) -> Path:
    """`A2` is the residual of `A4`; `B4` is the residual of `B2`."""
    return write_workbook(
        tmp_path / "a13_two_series.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "A2": "=A4",
                "B2": "=10",
                "A4": "=2",
                "B4": "=B2",
            },
        },
    )


def _two_series_bindings() -> dict:
    return bindings_document(
        _time_series("x", "Engine!A2:B2"),
        _time_series("y", "Engine!A4:B4"),
    )


def _qcraft_workbook(tmp_path: Path) -> Path:
    """Four-series Q-CRAFT canary: emp / prod / g identities rearrange by year."""
    return write_workbook(
        tmp_path / "a13_qcraft.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=200",
                "C2": "=B2*(1+C3/100)*(1+C4/100)",
                "A3": "=(A5/100-A4/100)/(1+A4/100)*100",
                "B3": "=3",
                "C3": "=3",
                "A4": "=2",
                "B4": "=(B5/100-B3/100)/(1+B3/100)*100",
                "C4": "=2",
                "A5": "=5",
                "B5": "=4",
                "C5": "=C2/B2*100-100",
            },
        },
    )


def _qcraft_bindings() -> dict:
    return bindings_document(
        _time_series("real_gdp_lcu", "Engine!A2:C2", output=False),
        _time_series("employment_growth", "Engine!A3:C3"),
        _time_series("labour_productivity_growth", "Engine!A4:C4"),
        _time_series("real_gdp_growth", "Engine!A5:C5"),
    )


def _qcraft_scc() -> tuple[str, ...]:
    return (
        "real_gdp_lcu",
        "employment_growth",
        "labour_productivity_growth",
        "real_gdp_growth",
    )


def test_identity_flip_is_not_a_cell_cycle(tmp_path: Path) -> None:
    workbook = _two_series_workbook(tmp_path)
    graph = create_dependency_graph(
        workbook,
        ["Engine!A2", "Engine!B2", "Engine!A4", "Engine!B4"],
        load_values=True,
    )
    assert graph.cycle_report().has_must_cycles is False


def test_identity_flip_emits_demand_driven_scan(tmp_path: Path) -> None:
    workbook = _two_series_workbook(tmp_path)
    modules = generate_inverted(workbook, _two_series_bindings())
    internals = modules["internals.py"]
    assert "eval_instance" in internals
    catalog, _deps, graph = inverted_graph_parts(workbook, _two_series_bindings())
    assert plan_fused_scc(("x", "y"), catalog=catalog, graph=graph) is None
    pkg = load_package(modules, tmp_path, name="a13_two")
    assert pkg.compute_x() == pytest.approx((2.0, 10.0))
    assert pkg.compute_y() == pytest.approx((2.0, 10.0))


def test_identity_flip_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _two_series_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _two_series_bindings()), tmp_path, name="a13_num"
    )
    graph = create_dependency_graph(
        workbook,
        ["Engine!A2", "Engine!B2", "Engine!A4", "Engine!B4"],
        load_values=True,
    )
    expected = FormulaEvaluator(graph).evaluate(
        ["Engine!A2", "Engine!B2", "Engine!A4", "Engine!B4"]
    )
    assert pkg.compute_x() == pytest.approx((expected["Engine!A2"], expected["Engine!B2"]))
    assert pkg.compute_y() == pytest.approx((expected["Engine!A4"], expected["Engine!B4"]))


def test_qcraft_identity_flip_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = _qcraft_workbook(tmp_path)
    targets = [
        "Engine!A2",
        "Engine!B2",
        "Engine!C2",
        "Engine!A3",
        "Engine!B3",
        "Engine!C3",
        "Engine!A4",
        "Engine!B4",
        "Engine!C4",
        "Engine!A5",
        "Engine!B5",
        "Engine!C5",
    ]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    assert graph.cycle_report().has_must_cycles is False
    modules = generate_inverted(workbook, _qcraft_bindings())
    internals = modules["internals.py"]
    assert "eval_instance" in internals
    catalog, _deps, graph_bound = inverted_graph_parts(workbook, _qcraft_bindings())
    assert plan_fused_scc(_qcraft_scc(), catalog=catalog, graph=graph_bound) is None
    pkg = load_package(modules, tmp_path, name="a13_qc")
    expected = FormulaEvaluator(graph).evaluate(targets)
    assert pkg.compute_employment_growth() == pytest.approx(
        (expected["Engine!A3"], expected["Engine!B3"], expected["Engine!C3"])
    )
    assert pkg.compute_labour_productivity_growth() == pytest.approx(
        (expected["Engine!A4"], expected["Engine!B4"], expected["Engine!C4"])
    )
    assert pkg.compute_real_gdp_growth() == pytest.approx(
        (expected["Engine!A5"], expected["Engine!B5"], expected["Engine!C5"])
    )
    assert pkg.internals.real_gdp_lcu() == pytest.approx(
        (expected["Engine!A2"], expected["Engine!B2"], expected["Engine!C2"])
    )


def test_same_column_cycle_still_fail_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a13_same_col.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": "=A4",
                "A4": "=A2",
            },
        },
    )
    document = bindings_document(
        _time_series("x", "Engine!A2"),
        _time_series("y", "Engine!A4"),
    )
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        generate_inverted(workbook, document)
