"""Access class `affine`: `f(i) = a*i + b` with integer `a != 1`."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.deps import (
    collect_all_dependence_edges,
    plan_indices,
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


def _time_series(series_id: str, data_range: str, *, output: bool = False) -> dict:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction="output" if output else "input",
        header_row=1,
    )


def _decimate_workbook(tmp_path: Path) -> Path:
    """Consumer `i` reads producer `2*i` (`A2`, `C2`, `E2`)."""
    return write_workbook(
        tmp_path / "affine_decimate.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "D1": 2012,
                "E1": 2013,
                "A2": 10.0,
                "B2": 20.0,
                "C2": 30.0,
                "D2": 40.0,
                "E2": 50.0,
                "A3": "=A2",
                "B3": "=C2",
                "C3": "=E2",
            },
        },
    )


def _decimate_bindings() -> dict:
    return bindings_document(
        _time_series("source", "Engine!A2:E2"),
        _time_series("sampled", "Engine!A3:C3", output=True),
    )


def _reverse_workbook(tmp_path: Path) -> Path:
    """Consumer `i` reads producer `n - 1 - i`."""
    return write_workbook(
        tmp_path / "affine_reverse.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "A2": 10.0,
                "B2": 20.0,
                "C2": 30.0,
                "A3": "=C2",
                "B3": "=B2",
                "C3": "=A2",
            },
        },
    )


def _reverse_bindings() -> dict:
    return bindings_document(
        _time_series("source", "Engine!A2:C2"),
        _time_series("reversed", "Engine!A3:C3", output=True),
    )


def _accesses(edges: tuple, consumer_id: str, producer_id: str) -> set[str]:
    return {
        edge.access
        for edge in edges
        if edge.consumer_id == consumer_id and edge.producer_id == producer_id
    }


def _affine_params(edges: tuple, consumer_id: str, producer_id: str) -> set[tuple[int, int]]:
    return {
        (edge.coeff, edge.offset)
        for edge in edges
        if edge.consumer_id == consumer_id and edge.producer_id == producer_id
    }


def test_decimate_edges_are_affine_not_mixed_identity_shift(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_decimate_workbook(tmp_path), _decimate_bindings())
    edges = collect_all_dependence_edges(catalog, graph)
    assert _accesses(edges, "sampled", "source") == {"affine"}
    assert _affine_params(edges, "sampled", "source") == {(2, 0)}
    sampled = deps["sampled"]
    assert sampled.affine_maps == {"source": (2, 0)}
    assert sampled.index_maps["source"] == (0, 2, 4)
    assert sampled.aligned_ids == frozenset({"source"})
    assert len(catalog.get("sampled").statements) == 1


def test_reverse_edges_are_affine(tmp_path: Path) -> None:
    catalog, deps, graph = inverted_graph_parts(_reverse_workbook(tmp_path), _reverse_bindings())
    edges = collect_all_dependence_edges(catalog, graph)
    assert _accesses(edges, "reversed", "source") == {"affine"}
    assert _affine_params(edges, "reversed", "source") == {(-1, 2)}
    assert deps["reversed"].affine_maps == {"source": (-1, 2)}
    assert deps["reversed"].index_maps["source"] == (2, 1, 0)
    assert len(catalog.get("reversed").statements) == 1


def test_plan_indices_maps_affine_image_without_index_map(tmp_path: Path) -> None:
    catalog, deps, _graph = inverted_graph_parts(_decimate_workbook(tmp_path), _decimate_bindings())
    deps["sampled"].index_maps = {}
    result, _call = plan_indices(catalog.get("sampled"), catalog=catalog, deps=deps)
    assert result["source"] == (0, 2, 4)


def test_decimate_emit_uses_strided_range_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = _decimate_workbook(tmp_path)
    modules = generate_inverted(workbook, _decimate_bindings())
    api = modules["api.py"]
    assert "take(source, range(0, 6, 2))" in api
    assert "take(source, (0, 2, 4))" not in api
    pkg = load_package(modules, tmp_path, name="affine_decimate")
    got = pkg.compute_sampled(source=(10.0, 20.0, 30.0, 40.0, 50.0))
    assert got == pytest.approx((10.0, 30.0, 50.0))
    graph = create_dependency_graph(
        workbook, ["Engine!A3", "Engine!B3", "Engine!C3"], load_values=True
    )
    expected = FormulaEvaluator(graph).evaluate(["Engine!A3", "Engine!B3", "Engine!C3"])
    assert got == pytest.approx(
        (expected["Engine!A3"], expected["Engine!B3"], expected["Engine!C3"])
    )


def test_reverse_emit_preserves_decreasing_order_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = _reverse_workbook(tmp_path)
    modules = generate_inverted(workbook, _reverse_bindings())
    api = modules["api.py"]
    assert "take(source, range(2, -1, -1))" in api
    pkg = load_package(modules, tmp_path, name="affine_reverse")
    got = pkg.compute_reversed(source=(10.0, 20.0, 30.0))
    assert got == pytest.approx((30.0, 20.0, 10.0))
    graph = create_dependency_graph(
        workbook, ["Engine!A3", "Engine!B3", "Engine!C3"], load_values=True
    )
    expected = FormulaEvaluator(graph).evaluate(["Engine!A3", "Engine!B3", "Engine!C3"])
    assert got == pytest.approx(
        (expected["Engine!A3"], expected["Engine!B3"], expected["Engine!C3"])
    )
