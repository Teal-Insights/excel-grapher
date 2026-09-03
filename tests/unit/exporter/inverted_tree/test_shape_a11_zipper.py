"""Layer A11 — period-lag zippers are a cell DAG and a series SCC (#603)."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.ast_emit import emit_rung2_scc, emit_rung3_scc
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.runtime import (
    XlError,
    as_measure,
    demand_instance,
    eval_instance,
    is_error,
    live_measure,
)
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _zipper_workbook(tmp_path: Path) -> Path:
    """Cell-acyclic zipper from `plans/inverted-tree-scheduling.md`.

    `debt_t = debt_{t-1} + adj_t` and `adj_t = debt_{t-1} * r` (lag only).
    """
    return write_workbook(
        tmp_path / "a11_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+B3",
                "C2": "=B2+C3",
                "B3": "=A2*0.02",
                "C3": "=B2*0.02",
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


def _vertical_zipper_workbook(tmp_path: Path) -> Path:
    """Same zipper laid out down a column. B2 reads B1 is a genuine t-1 lag."""
    return write_workbook(
        tmp_path / "a11_vertical_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "A3": 2011,
                "B1": "=100",
                "B2": "=B1+C2",
                "B3": "=B2+C3",
                "C2": "=B1*0.02",
                "C3": "=B2*0.02",
            },
        },
    )


def _vertical_zipper_bindings() -> dict:
    return bindings_document(
        series_entry(
            "debt",
            "Engine!B1:B3",
            layout="series",
            direction="output",
            label_column="A",
        ),
        series_entry(
            "adjustment",
            "Engine!C2:C3",
            layout="series",
            direction="internal",
            label_column="A",
        ),
    )


def _offset_zipper_workbook(tmp_path: Path) -> Path:
    """Adjustment sits in E:F; column subtraction invents a look-ahead."""
    return write_workbook(
        tmp_path / "a11_offset_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+E3",
                "C2": "=B2+F3",
                "E1": 2010,
                "F1": 2011,
                "E3": "=A2*0.02",
                "F3": "=B2*0.02",
            },
        },
    )


def _offset_zipper_bindings() -> dict:
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
            "Engine!E3:F3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def _cross_sheet_zipper_workbook(tmp_path: Path) -> Path:
    """Adjustment lives on another sheet; columns are not co-planar."""
    return write_workbook(
        tmp_path / "a11_cross_sheet_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+Helper!C2",
                "C2": "=B2+Helper!D2",
            },
            "Helper": {
                "C1": 2010,
                "D1": 2011,
                "C2": "=Engine!A2*0.02",
                "D2": "=Engine!B2*0.02",
            },
        },
    )


def _cross_sheet_zipper_bindings() -> dict:
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
            "Helper!C2:D2",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def _vertical_scan_workbook(tmp_path: Path) -> Path:
    """Single-series vertical scan; predecessor_address already uses expansion order."""
    return write_workbook(
        tmp_path / "a11_vertical_scan.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "A3": 2011,
                "B1": "=100",
                "B2": "=B1*1.02",
                "B3": "=B2*1.02",
            },
        },
    )


def _vertical_scan_bindings() -> dict:
    return bindings_document(
        series_entry(
            "debt",
            "Engine!B1:B3",
            layout="series",
            direction="output",
            label_column="A",
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


def _vertical_simultaneous_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a11_simultaneous_vertical.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "A3": 2011,
                "B1": "=100",
                "B2": "=B1+C2",
                "B3": "=B2+C3",
                "C2": "=B2-B1",
                "C3": "=B3-B2",
            },
        },
    )


def _zipper_orientation(orientation: str) -> tuple:
    if orientation == "horizontal":
        return (
            _zipper_workbook,
            _zipper_bindings,
            ["Engine!A2", "Engine!B2", "Engine!C2"],
            ["Engine!B3", "Engine!C3"],
        )
    return (
        _vertical_zipper_workbook,
        _vertical_zipper_bindings,
        ["Engine!B1", "Engine!B2", "Engine!B3"],
        ["Engine!C2", "Engine!C3"],
    )


@pytest.mark.parametrize("orientation", ["horizontal", "vertical"])
def test_lag_zipper_emits_fused_union_loop(tmp_path: Path, orientation: str) -> None:
    workbook_fn, bindings_fn, _debt, _adj = _zipper_orientation(orientation)
    workbook = workbook_fn(tmp_path)
    modules = generate_inverted(workbook, bindings_fn())
    internals = modules["internals.py"]
    api = modules["api.py"]
    assert "cyclic formula-series" not in internals
    assert "eval_instance" not in internals
    assert "for t in range(" in internals
    assert internals.count("for t in range(") == 1
    pkg = load_package(modules, tmp_path, name=f"a11_zip_{orientation[:1]}")
    got = pkg.compute_debt()
    assert got == pytest.approx((100.0, 102.0, 104.04))
    assert "scan_debt_adjustment" in internals
    assert "internals.scan_debt_adjustment" in api


@pytest.mark.parametrize("orientation", ["horizontal", "vertical"])
def test_lag_zipper_matches_formula_evaluator(tmp_path: Path, orientation: str) -> None:
    workbook_fn, bindings_fn, debt_cells, adj_cells = _zipper_orientation(orientation)
    workbook = workbook_fn(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, bindings_fn()), tmp_path, name=f"a11_num_{orientation[:1]}"
    )
    targets = [*debt_cells, *adj_cells]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    assert graph.cycle_report().has_must_cycles is False
    expected = FormulaEvaluator(graph).evaluate(targets)
    got = pkg.compute_debt()
    assert got == pytest.approx(tuple(expected[cell] for cell in debt_cells))
    adj = pkg.internals.adjustment()
    assert adj == pytest.approx(tuple(expected[cell] for cell in adj_cells))


@pytest.mark.parametrize(
    ("workbook_fn", "bindings_fn"),
    [
        (_simultaneous_workbook, _zipper_bindings),
        (_vertical_simultaneous_workbook, _vertical_zipper_bindings),
    ],
    ids=["horizontal", "vertical"],
)
def test_same_year_cell_cycle_still_fail_closed(tmp_path: Path, workbook_fn, bindings_fn) -> None:
    workbook = workbook_fn(tmp_path)
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        generate_inverted(workbook, bindings_fn())


def _exec_scan(body: list[str], names: set[str]) -> tuple[tuple[object, ...], ...]:
    runtime = {
        "XlError": XlError,
        "as_measure": as_measure,
        "demand_instance": demand_instance,
        "eval_instance": eval_instance,
        "is_error": is_error,
        "live_measure": live_measure,
    }
    ns: dict[str, Any] = {name: runtime[name] for name in names if name in runtime}
    exec("def scan():\n" + "\n".join(body), ns)
    scan = cast(Callable[[], tuple[tuple[object, ...], ...]], ns["scan"])
    return scan()


@pytest.mark.parametrize("orientation", ["horizontal", "vertical"])
def test_fused_loop_agrees_with_rung3_oracle(tmp_path: Path, orientation: str) -> None:
    workbook_fn, bindings_fn, _debt, _adj = _zipper_orientation(orientation)
    catalog, deps, graph = inverted_graph_parts(workbook_fn(tmp_path), bindings_fn())
    scc = ("debt", "adjustment")
    fused, fused_used = emit_rung2_scc(scc, catalog=catalog, deps=deps, graph=graph)
    demand, demand_used = emit_rung3_scc(scc, catalog=catalog, deps=deps, graph=graph)
    assert _exec_scan(fused, fused_used) == _exec_scan(demand, demand_used)


def test_offset_helper_block_stays_on_rung2(tmp_path: Path) -> None:
    workbook = _offset_zipper_workbook(tmp_path)
    modules = generate_inverted(workbook, _offset_zipper_bindings())
    internals = modules["internals.py"]
    assert "eval_instance" not in internals
    assert "for t in range(" in internals
    pkg = load_package(modules, tmp_path, name="a11_off")
    assert pkg.compute_debt() == pytest.approx((100.0, 102.0, 104.04))


def test_cross_sheet_zipper_joins_on_time_period(tmp_path: Path) -> None:
    workbook = _cross_sheet_zipper_workbook(tmp_path)
    modules = generate_inverted(workbook, _cross_sheet_zipper_bindings())
    internals = modules["internals.py"]
    assert "eval_instance" not in internals
    assert "for t in range(" in internals
    pkg = load_package(modules, tmp_path, name="a11_xsheet")
    assert pkg.compute_debt() == pytest.approx((100.0, 102.0, 104.04))


def test_vertical_scan_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _vertical_scan_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _vertical_scan_bindings()), tmp_path, name="a11_vscan"
    )
    graph = create_dependency_graph(
        workbook, ["Engine!B1", "Engine!B2", "Engine!B3"], load_values=True
    )
    expected = FormulaEvaluator(graph).evaluate(["Engine!B1", "Engine!B2", "Engine!B3"])
    assert pkg.compute_debt() == pytest.approx(
        (expected["Engine!B1"], expected["Engine!B2"], expected["Engine!B3"])
    )
