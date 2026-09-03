"""Layer A13 — residual identity can flip across years without a cell cycle."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import FusedRegion, plan_fused_scc
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


def _vertical_two_series_workbook(tmp_path: Path) -> Path:
    """`B1` is the residual of `C1`; `C2` is the residual of `B2`."""
    return write_workbook(
        tmp_path / "a13_two_series_vertical.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "B1": "=C1",
                "B2": "=10",
                "C1": "=2",
                "C2": "=B2",
            },
        },
    )


def _vertical_two_series_bindings() -> dict:
    return bindings_document(
        series_entry(
            "x",
            "Engine!B1:B2",
            layout="series",
            direction="output",
            label_column="A",
        ),
        series_entry(
            "y",
            "Engine!C1:C2",
            layout="series",
            direction="output",
            label_column="A",
        ),
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


@pytest.mark.parametrize(
    ("workbook_fn", "bindings_fn", "x_cells", "y_cells", "pkg_name"),
    [
        (
            _two_series_workbook,
            _two_series_bindings,
            ["Engine!A2", "Engine!B2"],
            ["Engine!A4", "Engine!B4"],
            "a13_h",
        ),
        (
            _vertical_two_series_workbook,
            _vertical_two_series_bindings,
            ["Engine!B1", "Engine!B2"],
            ["Engine!C1", "Engine!C2"],
            "a13_v",
        ),
    ],
    ids=["horizontal", "vertical"],
)
def test_identity_flip_emits_region_local_fused_scan(
    tmp_path: Path,
    workbook_fn,
    bindings_fn,
    x_cells: list[str],
    y_cells: list[str],
    pkg_name: str,
) -> None:
    workbook = workbook_fn(tmp_path)
    modules = generate_inverted(workbook, bindings_fn())
    internals = modules["internals.py"]
    assert "eval_instance" not in internals
    assert "for t in range(" in internals
    assert "if t == 0:" in internals
    assert "elif t == 1:" in internals
    catalog, _deps, graph = inverted_graph_parts(workbook, bindings_fn())
    plan = plan_fused_scc(("x", "y"), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan.regions == (
        FusedRegion(start=0, stop=1, body_order=("y", "x")),
        FusedRegion(start=1, stop=2, body_order=("x", "y")),
    )
    pkg = load_package(modules, tmp_path, name=pkg_name)
    assert pkg.compute_x() == pytest.approx((2.0, 10.0))
    assert pkg.compute_y() == pytest.approx((2.0, 10.0))


@pytest.mark.parametrize(
    ("workbook_fn", "bindings_fn", "x_cells", "y_cells", "pkg_name"),
    [
        (
            _two_series_workbook,
            _two_series_bindings,
            ["Engine!A2", "Engine!B2"],
            ["Engine!A4", "Engine!B4"],
            "a13_num_h",
        ),
        (
            _vertical_two_series_workbook,
            _vertical_two_series_bindings,
            ["Engine!B1", "Engine!B2"],
            ["Engine!C1", "Engine!C2"],
            "a13_num_v",
        ),
    ],
    ids=["horizontal", "vertical"],
)
def test_identity_flip_matches_formula_evaluator(
    tmp_path: Path,
    workbook_fn,
    bindings_fn,
    x_cells: list[str],
    y_cells: list[str],
    pkg_name: str,
) -> None:
    workbook = workbook_fn(tmp_path)
    pkg = load_package(generate_inverted(workbook, bindings_fn()), tmp_path, name=pkg_name)
    targets = [*x_cells, *y_cells]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    expected = FormulaEvaluator(graph).evaluate(targets)
    assert pkg.compute_x() == pytest.approx(tuple(expected[cell] for cell in x_cells))
    assert pkg.compute_y() == pytest.approx(tuple(expected[cell] for cell in y_cells))


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
    assert "eval_instance" not in internals
    assert "for t in range(" in internals
    catalog, _deps, graph_bound = inverted_graph_parts(workbook, _qcraft_bindings())
    plan = plan_fused_scc(_qcraft_scc(), catalog=catalog, graph=graph_bound)
    assert plan is not None
    assert plan.regions == (
        FusedRegion(
            start=0,
            stop=1,
            body_order=(
                "real_gdp_lcu",
                "labour_productivity_growth",
                "real_gdp_growth",
                "employment_growth",
            ),
        ),
        FusedRegion(
            start=1,
            stop=2,
            body_order=(
                "real_gdp_lcu",
                "employment_growth",
                "real_gdp_growth",
                "labour_productivity_growth",
            ),
        ),
        FusedRegion(
            start=2,
            stop=3,
            body_order=(
                "employment_growth",
                "labour_productivity_growth",
                "real_gdp_lcu",
                "real_gdp_growth",
            ),
        ),
    )
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


def test_look_ahead_stays_on_rung3(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a13_lookahead.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "A2": "=B4",
                "B2": "=10",
                "A4": "=1",
                "B4": "=B2",
            },
        },
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, _two_series_bindings())
    assert plan_fused_scc(("x", "y"), catalog=catalog, graph=graph) is None
    internals = generate_inverted(workbook, _two_series_bindings())["internals.py"]
    assert "eval_instance" in internals


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
