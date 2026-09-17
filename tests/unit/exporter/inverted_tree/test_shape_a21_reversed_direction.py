"""Layer A21 — reversed loop direction for all-negative-distance SCCs (#614).

An SCC whose intra-SCC distances are all <= 0 is fusible with the loop run in
reverse (Allen-Kennedy loop direction selection). Single-series backward
recurrence is a rung-1 reversed scan.
"""

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

# ---------------------------------------------------------------------------
# Case 1: Terminal-value backward recursion (single-series scan, Rung 1)
# ---------------------------------------------------------------------------


def _horizontal_terminal_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a21_terminal_h.xlsx",
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
        tmp_path / "a21_terminal_v.xlsx",
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
            "a21_term_h",
        ),
        (
            _vertical_terminal_workbook,
            _vertical_terminal_bindings,
            ["Engine!B1", "Engine!B2", "Engine!B3"],
            "a21_term_v",
        ),
    ],
    ids=["horizontal", "vertical"],
)
def test_terminal_backward_recursion_emits_reversed_scan_and_matches_evaluator(
    tmp_path: Path,
    workbook_fn,
    bindings_fn,
    cells: list[str],
    pkg_name: str,
) -> None:
    workbook = workbook_fn(tmp_path)
    doc = bindings_fn()
    catalog, deps, graph = inverted_graph_parts(workbook, doc)
    catalog.get("value")
    assert deps["value"].is_scan is True
    assert deps["value"].scan_direction == "reversed"

    modules = generate_inverted(workbook, doc)
    pkg = load_package(modules, tmp_path, name=pkg_name)
    graph_full = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph_full).evaluate(cells)
    got = pkg.compute_value()
    assert [value for _, value in got.items()] == pytest.approx(tuple(expected[c] for c in cells))
    assert [value for _, value in got.items()] == pytest.approx((81.0, 90.0, 100.0))


# ---------------------------------------------------------------------------
# Case 2: Look-ahead zipper (multi-series SCC, Rung 2)
# ---------------------------------------------------------------------------


def _lookahead_zipper_workbook(tmp_path: Path) -> Path:
    """Look-ahead zipper: value_t = value_{t+1}*0.9 + flow_t, flow_t = value_{t+1}*0.01."""
    return write_workbook(
        tmp_path / "a21_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=B2*0.9+A3",
                "B2": "=C2*0.9+B3",
                "C2": "=100",
                "A3": "=B2*0.01",
                "B3": "=C2*0.01",
            },
        },
    )


def _lookahead_zipper_bindings() -> dict:
    return bindings_document(
        series_entry(
            "value",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        series_entry(
            "flow",
            "Engine!A3:B3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def _vertical_lookahead_zipper_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a21_vertical_zipper.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "A2": 2010,
                "A3": 2011,
                "B1": "=B2*0.9+C1",
                "B2": "=B3*0.9+C2",
                "B3": "=100",
                "C1": "=B2*0.01",
                "C2": "=B3*0.01",
            },
        },
    )


def _vertical_lookahead_zipper_bindings() -> dict:
    return bindings_document(
        series_entry(
            "value",
            "Engine!B1:B3",
            layout="series",
            direction="output",
            label_column="A",
        ),
        series_entry(
            "flow",
            "Engine!C1:C2",
            layout="series",
            direction="internal",
            label_column="A",
        ),
    )


@pytest.mark.parametrize(
    ("workbook_fn", "bindings_fn", "value_cells", "flow_cells", "pkg_name"),
    [
        (
            _lookahead_zipper_workbook,
            _lookahead_zipper_bindings,
            ["Engine!A2", "Engine!B2", "Engine!C2"],
            ["Engine!A3", "Engine!B3"],
            "a21_zip_h",
        ),
        (
            _vertical_lookahead_zipper_workbook,
            _vertical_lookahead_zipper_bindings,
            ["Engine!B1", "Engine!B2", "Engine!B3"],
            ["Engine!C1", "Engine!C2"],
            "a21_zip_v",
        ),
    ],
    ids=["horizontal", "vertical"],
)
def test_lookahead_zipper_emits_fused_reversed_loop_and_matches_evaluator(
    tmp_path: Path,
    workbook_fn,
    bindings_fn,
    value_cells: list[str],
    flow_cells: list[str],
    pkg_name: str,
) -> None:
    workbook = workbook_fn(tmp_path)
    doc = bindings_fn()
    pkg = load_package(generate_inverted(workbook, doc), tmp_path, name=pkg_name)
    all_cells = [*value_cells, *flow_cells]
    graph_full = create_dependency_graph(workbook, all_cells, load_values=True)
    expected = FormulaEvaluator(graph_full).evaluate(all_cells)
    got_value = pkg.compute_value()
    assert [value for _, value in got_value.items()] == pytest.approx(
        tuple(expected[c] for c in value_cells)
    )
    got_flow = pkg.internals.scan_value().flow
    assert [value for _, value in got_flow.items()] == pytest.approx(
        tuple(expected[c] for c in flow_cells)
    )


# ---------------------------------------------------------------------------
# Case 3: Descending-year layout (latest year leftmost)
# ---------------------------------------------------------------------------


def _descending_year_zipper_workbook(tmp_path: Path) -> Path:
    """A1=2011, B1=2010, C1=2009. debt_t reads cell to the right (t-1)."""
    return write_workbook(
        tmp_path / "a21_desc_zipper.xlsx",
        {
            "Engine": {
                "A1": 2011,
                "B1": 2010,
                "C1": 2009,
                "A2": "=B2*1.02+A3",
                "B2": "=C2*1.02+B3",
                "C2": "=100",
                "A3": "=B2*0.02",
                "B3": "=C2*0.02",
            },
        },
    )


def _descending_year_zipper_bindings() -> dict:
    return bindings_document(
        series_entry(
            "debt",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        series_entry(
            "adj",
            "Engine!A3:B3",
            layout="series",
            direction="internal",
            header_row=1,
        ),
    )


def test_descending_year_layout_fuses_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = _descending_year_zipper_workbook(tmp_path)
    doc = _descending_year_zipper_bindings()
    pkg = load_package(generate_inverted(workbook, doc), tmp_path, name="a21_desc_zip")
    cells = ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!A3", "Engine!B3"]
    graph_full = create_dependency_graph(workbook, cells, load_values=True)
    expected = FormulaEvaluator(graph_full).evaluate(cells)
    got = pkg.compute_debt()
    assert [value for _, value in got.items()] == pytest.approx(
        tuple(expected[f"Engine!{col}2"] for col in ("A", "B", "C"))
    )


# ---------------------------------------------------------------------------
# Refusal: mixed signs drop to Rung 3, same-index cycles fail closed
# ---------------------------------------------------------------------------


def test_mixed_signs_refuses_fused_plan(tmp_path: Path) -> None:
    """An SCC with both forward and look-ahead edges cannot fuse in either direction."""
    workbook = write_workbook(
        tmp_path / "mixed_scc.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": "=100",
                "B2": "=A2+C3",  # reads t-1 from debt, and t+1 from other
                "C2": "=B2+10",
                "A3": "=B2*0.1",  # reads t+1 from debt
                "B3": "=A2*0.1",  # reads t-1 from debt
                "C3": "=10",
            },
        },
    )
    doc = bindings_document(
        series_entry("x", "Engine!A2:C2", layout="series", direction="output", header_row=1),
        series_entry("y", "Engine!A3:C3", layout="series", direction="internal", header_row=1),
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, doc)
    pkg = load_package(generate_inverted(workbook, doc), tmp_path, name="a21_mixed")
    cells = ["Engine!A2", "Engine!B2", "Engine!C2", "Engine!A3", "Engine!B3", "Engine!C3"]
    expected = FormulaEvaluator(graph).evaluate(cells)
    got = pkg.compute_x()
    assert [value for _, value in got.items()] == pytest.approx(
        tuple(expected[f"Engine!{col}2"] for col in ("A", "B", "C"))
    )
