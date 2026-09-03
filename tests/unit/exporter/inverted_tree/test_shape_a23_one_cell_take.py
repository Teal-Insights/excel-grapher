"""Layer A23 — 1-cell TIME_PERIOD hosts honour the aligned `take` window (#626).

A keyed 1-cell series is `is_scalar` (helper `index_var is None`) but still
joins longer producers on `TIME_PERIOD`. `_aligned_call_arg` takes the producer
to the matching catalog slot; the helper body must index that window, not the
producer catalog.
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


def _time_series(
    series_id: str, data_range: str, *, output: bool = False, internal: bool = False
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


def _one_cell_workbook(tmp_path: Path) -> Path:
    """`growth` is `A:C` (3 years). `last_growth` is the 2011 cell only (`=C2`)."""
    return write_workbook(
        tmp_path / "a23_one_cell.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": 3.0,
                "B2": 4.0,
                "C2": 5.0,
                "C3": "=C2",
            },
        },
    )


def _one_cell_bindings() -> dict:
    return bindings_document(
        _time_series("growth", "Engine!A2:C2"),
        _time_series("last_growth", "Engine!C3:C3", output=True),
    )


def _shared_workbook(tmp_path: Path) -> Path:
    """Same 1-cell read plus a full-span output so the runner keeps all of `growth`."""
    return write_workbook(
        tmp_path / "a23_shared.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": 3.0,
                "B2": 4.0,
                "C2": 5.0,
                "A4": "=A2",
                "B4": "=B2",
                "C4": "=C2",
                "C3": "=C2",
            },
        },
    )


def _shared_bindings() -> dict:
    return bindings_document(
        _time_series("growth", "Engine!A2:C2"),
        _time_series("last_growth", "Engine!C3:C3", output=True),
        _time_series("all_growth", "Engine!A4:C4", output=True),
    )


def _scalar_host_bindings() -> dict:
    return bindings_document(
        _time_series("growth", "Engine!A2:C2"),
        series_entry("last_growth", "Engine!C3", layout="scalar", direction="output"),
    )


def test_one_cell_host_index_map_is_producer_slot(tmp_path: Path) -> None:
    catalog, deps, _graph = inverted_graph_parts(_one_cell_workbook(tmp_path), _one_cell_bindings())
    assert catalog.get("last_growth").is_scalar
    assert deps["last_growth"].index_maps["growth"] == (2,)
    assert "growth" in deps["last_growth"].aligned_ids


def test_one_cell_helper_indexes_taken_window(tmp_path: Path) -> None:
    workbook = _one_cell_workbook(tmp_path)
    modules = generate_inverted(workbook, _one_cell_bindings())
    api = modules["api.py"]
    internals = modules["internals.py"]
    assert "take(growth, (2,))" in api
    assert "growth[2]" not in internals
    assert "growth[0]" in internals
    pkg = load_package(modules, tmp_path, name="a23_one")
    got = pkg.compute_last_growth(growth=(3.0, 4.0, 5.0))
    if isinstance(got, tuple):
        assert len(got) == 1
        got = got[0]
    assert got == pytest.approx(5.0)


def test_one_cell_take_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _one_cell_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _one_cell_bindings()), tmp_path, name="a23_eval")
    graph = create_dependency_graph(workbook, ["Engine!C3"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Engine!C3"])
    got = pkg.compute_last_growth(growth=(3.0, 4.0, 5.0))
    if isinstance(got, tuple):
        got = got[0]
    assert got == pytest.approx(expected["Engine!C3"])


def test_shared_runner_takes_at_one_cell_call_site(tmp_path: Path) -> None:
    workbook = _shared_workbook(tmp_path)
    modules = generate_inverted(workbook, _shared_bindings())
    api = modules["api.py"]
    internals = modules["internals.py"]
    assert "last_growth(take(growth, (2,)))" in api
    assert "growth[2]" not in internals
    assert "growth[0]" in internals
    pkg = load_package(modules, tmp_path, name="a23_shared")
    last = pkg.compute_last_growth(growth=(3.0, 4.0, 5.0))
    if isinstance(last, tuple):
        last = last[0]
    assert last == pytest.approx(5.0)
    assert pkg.compute_all_growth(growth=(3.0, 4.0, 5.0)) == pytest.approx((3.0, 4.0, 5.0))


def test_scalar_host_keeps_catalog_subscript(tmp_path: Path) -> None:
    workbook = _one_cell_workbook(tmp_path)
    modules = generate_inverted(workbook, _scalar_host_bindings())
    api = modules["api.py"]
    internals = modules["internals.py"]
    assert "take(growth" not in api
    assert "growth[2]" in internals
    pkg = load_package(modules, tmp_path, name="a23_scalar")
    got = pkg.compute_last_growth(growth=(3.0, 4.0, 5.0))
    if isinstance(got, tuple):
        got = got[0]
    assert got == pytest.approx(5.0)
