"""Projected inverted-tree packages import and match FormulaEvaluator."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter import CodeGenerator, IdentityTransitCompression, OptimalCompression
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    load_package,
    series_entry,
    write_workbook,
)

_TARGETS = ["Engine!C6", "Outputs!B12", "Outputs!B14"]


def _identity_workbook(path: Path) -> Path:
    return write_workbook(
        path,
        {
            "Engine": {"C6": 10},
            "Outputs": {"B12": "=Engine!C6", "B14": "=Outputs!B12+1"},
        },
    )


def _bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Engine!C6", layout="scalar", direction="input"),
        series_entry("mirror", "Outputs!B12", layout="scalar", direction="internal"),
        series_entry(
            "next",
            "Outputs!B14",
            layout="scalar",
            direction="output",
            compute_name="compute_next",
        ),
    )


def _export_projected(workbook: Path, projection: object, tmp_path: Path, name: str):
    document = _bindings()
    bindings = validate_bindings_document(document)
    with CodeGenerator(projection) as gen:
        modules = gen.generate_modules(series_bindings=bindings, bindings_workbook=workbook)
    return load_package(modules, tmp_path, name=name)


def test_identity_projected_generate_modules_matches_evaluator(tmp_path: Path) -> None:
    workbook = _identity_workbook(tmp_path / "identity.xlsx")
    graph = create_dependency_graph(
        workbook,
        _TARGETS,
        load_values=True,
        capture_dependency_provenance=True,
    )
    projection = IdentityTransitCompression().project(graph)
    pkg = _export_projected(workbook, projection, tmp_path, "projected_inv")
    assert pkg.compute_next(seed=10.0) == (11.0,)
    assert pkg.compute_next(seed=7.0) == (8.0,)
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate(["Outputs!B14"])["Outputs!B14"] == 11.0


def test_optimal_projected_generate_modules_matches_evaluator(tmp_path: Path) -> None:
    workbook = _identity_workbook(tmp_path / "optimal.xlsx")
    graph = create_dependency_graph(
        workbook,
        _TARGETS,
        load_values=True,
        capture_dependency_provenance=True,
    )
    projection = OptimalCompression().project(graph)
    pkg = _export_projected(workbook, projection, tmp_path, "optimal_inv")
    assert pkg.compute_next(seed=10.0) == (11.0,)
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate(["Outputs!B14"])["Outputs!B14"] == 11.0
