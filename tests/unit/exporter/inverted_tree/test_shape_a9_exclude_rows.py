"""Layer A9 — catalog applies `exclude_rows` / `exclude_columns` (#600)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import resolve_series_binding, validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)


def _risks_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a9_risks.xlsx",
        {
            "Risks": {
                "B1": 2020,
                "C1": 2021,
                "A2": "Band",
                "B2": 1.0,
                "C2": 2.0,
                "A3": "Band",
                "B3": 3.0,
                "C3": 4.0,
            },
            "Outputs": {
                "A1": "=Risks!B2",
            },
        },
    )


def _shock_series(
    series_id: str,
    *,
    exclude_rows: list[int] | None = None,
    exclude_columns: list[str] | None = None,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "Risks",
        "data_range": "Risks!B2:C3",
        "layout": "matrix",
        "input": {"setter": {"name": f"set_{series_id}"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "SCENARIO",
                    "concept": "SCENARIO",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "row_label",
                        "label_column": "A",
                        "fill": True,
                        "read": "string",
                    },
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "column_header",
                        "header_row": 1,
                        "read": "int",
                    },
                },
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    if exclude_rows is not None:
        entry["exclude_rows"] = exclude_rows
    if exclude_columns is not None:
        entry["exclude_columns"] = exclude_columns
    return entry


def _interleaved_row_bindings() -> dict[str, Any]:
    return bindings_document(
        _shock_series("revenue_shocks", exclude_rows=[3]),
        _shock_series("expenditure_shocks", exclude_rows=[2]),
        series_entry("output_cell", "Outputs!A1", layout="scalar", direction="output"),
    )


def _measure(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def test_catalog_partitions_complementary_exclude_rows(tmp_path: Path) -> None:
    workbook = _risks_workbook(tmp_path)
    document = _interleaved_row_bindings()
    bindings = validate_bindings_document(document)
    catalog = build_catalog(bindings, workbook=workbook)
    revenue = catalog.get("revenue_shocks")
    expenditure = catalog.get("expenditure_shocks")
    assert revenue.cells == ("Risks!B2", "Risks!C2")
    assert expenditure.cells == ("Risks!B3", "Risks!C3")
    assert set(revenue.cells).isdisjoint(expenditure.cells)
    graph = create_dependency_graph(
        workbook,
        ["Risks!B2", "Risks!C2", "Risks!B3", "Risks!C3", "Outputs!A1"],
        load_values=True,
    )
    for series_id, bound in (("revenue_shocks", revenue), ("expenditure_shocks", expenditure)):
        series = next(s for s in bindings["series"] if s["id"] == series_id)
        resolved = resolve_series_binding(graph, workbook, series)
        assert resolved["ok"]
        assert bound.cells == tuple(leaf["address"] for leaf in resolved["leaves"])


def test_interleaved_matrices_emit_and_match_evaluator(tmp_path: Path) -> None:
    workbook = _risks_workbook(tmp_path)
    modules = generate_inverted(workbook, _interleaved_row_bindings())
    assert "ctx" not in modules["api.py"]
    pkg = load_package(modules, tmp_path, name="a9_rows")
    assert required_param_names(pkg.compute_output_cell) == ("revenue_shocks",)
    assert "expenditure_shocks" not in all_param_names(pkg.compute_output_cell)
    got = _measure(pkg.compute_output_cell(revenue_shocks=(1.0, 2.0)))
    graph = create_dependency_graph(workbook, ["Outputs!A1"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])
    assert got == pytest.approx(expected["Outputs!A1"])
    assert got == pytest.approx(1.0)


def test_catalog_partitions_complementary_exclude_columns(tmp_path: Path) -> None:
    workbook = _risks_workbook(tmp_path)
    bindings = validate_bindings_document(
        bindings_document(
            _shock_series("revenue_shocks", exclude_columns=["C"]),
            _shock_series("expenditure_shocks", exclude_columns=["B"]),
            series_entry("output_cell", "Outputs!A1", layout="scalar", direction="output"),
        )
    )
    catalog = build_catalog(bindings, workbook=workbook)
    assert catalog.get("revenue_shocks").cells == ("Risks!B2", "Risks!B3")
    assert catalog.get("expenditure_shocks").cells == ("Risks!C2", "Risks!C3")


def test_unfiltered_overlap_still_fail_closed(tmp_path: Path) -> None:
    workbook = _risks_workbook(tmp_path)
    document = bindings_document(
        _shock_series("revenue_shocks"),
        _shock_series("expenditure_shocks"),
        series_entry("output_cell", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match="bound to both"):
        generate_inverted(workbook, document)
