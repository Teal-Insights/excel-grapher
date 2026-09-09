"""Issue 708 — double-bind retained graph leaves to input/constant series.

A graph_leaf hole inside an internal/output series may also be named by one
input or constant series. Catalog ownership stays with the formula series;
the claimant supplies the value at emit time.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.deps import collect_series_edges
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings import validate_bindings_document, validate_series_bindings
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    load_package,
    series_entry,
)
from tests.unit.exporter.inverted_tree.test_partial_graph_overlap import (
    _emit_from_outputs,
    _output_graph,
    _series_leaf_workbook,
)


def _bound_leaf_document(
    data_range: str,
    *extra: dict[str, Any],
    seed_direction: str = "constant",
) -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "rate",
            "Inputs!B2:D2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry("seed", "Engine!C2", layout="scalar", direction=seed_direction),
        series_entry(
            "engine_row",
            data_range,
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
        *extra,
    )


def test_catalog_bound_leaf_constant_pairing(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = _bound_leaf_document(data_range)
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)

    series = catalog.get("engine_row")
    assert series.cells == ("Engine!B2", "Engine!C2", "Engine!D2")
    assert catalog.series_id_for("Engine!C2") == "engine_row"
    assert catalog.get("seed").cells == ("Engine!C2",)
    hole = series.hole_at(1)
    assert hole is not None
    assert hole.kind == "bound_leaf"
    assert hole.address == "Engine!C2"
    assert hole.claimant_id == "seed"

    edges = collect_series_edges(series, catalog=catalog, graph=graph)
    assert any(edge.consumer_id == "engine_row" and edge.producer_id == "seed" for edge in edges)


def test_catalog_bound_leaf_input_pairing(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = _bound_leaf_document(data_range, seed_direction="input")
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    hole = catalog.get("engine_row").hole_at(1)
    assert hole is not None
    assert hole.kind == "bound_leaf"
    assert hole.claimant_id == "seed"
    assert catalog.series_id_for("Engine!C2") == "engine_row"


def test_catalog_third_series_on_bound_leaf_raises(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = _bound_leaf_document(
        data_range,
        series_entry("seed_again", "Engine!C2", layout="scalar", direction="input"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.raises(InvertedTreeExportError, match="bound to both"):
        build_catalog(bindings, workbook=workbook, graph=graph)


def test_catalog_double_binding_formula_cell_raises(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = bindings_document(
        series_entry(
            "rate",
            "Inputs!B2:D2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry("hijack", "Engine!B2", layout="scalar", direction="constant"),
        series_entry(
            "engine_row",
            data_range,
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.raises(InvertedTreeExportError, match="bound to both"):
        build_catalog(bindings, workbook=workbook, graph=graph)


def test_validate_accepts_bound_leaf_and_drops_leaf_notice(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = _bound_leaf_document(data_range)
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is True
    engine_issues = [i for i in report["issues"] if i.get("series_id") == "engine_row"]
    assert not any(i["code"] == "leaf_in_formula_series" for i in engine_issues)
    assert not any(i["level"] == "error" for i in report["issues"])


def test_validate_rejects_formula_cell_claimant(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = bindings_document(
        series_entry(
            "rate",
            "Inputs!B2:D2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry("hijack", "Engine!B2", layout="scalar", direction="constant"),
        series_entry(
            "engine_row",
            data_range,
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is False
    assert any("bound to both" in i["message"] for i in report["issues"] if i["level"] == "error")


def test_emit_bound_leaf_constant_is_parameter(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = _bound_leaf_document(data_range)
    graph = _output_graph(workbook, document)
    modules = _emit_from_outputs(workbook, document)
    internals = modules["internals.py"]
    assert "seed:" in internals
    assert "def engine_row(" in internals

    pkg = load_package(modules, tmp_path, name="bound_leaf_constant")
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])
    rate = pkg.data.Rate.from_records(
        domain=pkg.data.RATE_REQUIRED,
        records=(((2021,), 1.0), ((2023,), 3.0)),
    )
    assert pkg.compute_result(rate=rate) == pytest.approx(expected["Outputs!A1"])
    with pkg.data.overrides(SEED=99.0):
        assert pkg.compute_result(rate=rate) == pytest.approx(2.0 + 99.0 + 6.0)
    assert "seed" in pkg.compute_result.__constants__
