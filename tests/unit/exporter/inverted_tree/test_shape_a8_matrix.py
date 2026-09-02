"""Layer A8 — `layout: matrix` is a 1-D sequence in canonical cell order (#599)."""

from __future__ import annotations

import inspect
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import _layout_of, build_catalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)


def _profile_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a8_profile.xlsx",
        {
            "Profile": {
                "B1": 2020,
                "C1": 2021,
                "A2": "France",
                "B2": 10.0,
                "C2": 11.0,
                "A3": "Kenya",
                "B3": 20.0,
                "C3": 21.0,
            },
            "Outputs": {
                "A1": "=Profile!B2",
            },
        },
    )


def _profile_table_series() -> dict:
    return {
        "id": "profile_table",
        "sheet": "Profile",
        "data_range": "Profile!B2:C3",
        "layout": "matrix",
        "constant": {},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "COUNTRY",
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "row_label",
                        "label_column": "A",
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
        "key": ["COUNTRY", "TIME_PERIOD"],
    }


def _profile_bindings() -> dict:
    return bindings_document(
        _profile_table_series(),
        series_entry(
            "output_cell",
            "Outputs!A1",
            layout="scalar",
            direction="output",
        ),
    )


def _measure(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def test_catalog_accepts_matrix_as_row_major_sequence(tmp_path: Path) -> None:
    workbook = _profile_workbook(tmp_path)
    catalog = build_catalog(validate_bindings_document(_profile_bindings()), workbook=workbook)
    series = catalog.get("profile_table")
    assert series.layout == "matrix"
    assert series.cells == ("Profile!B2", "Profile!C2", "Profile!B3", "Profile!C3")
    assert series.is_sequence
    assert not series.is_scalar
    assert not series.is_time_series
    assert series.key_fields == ("COUNTRY", "TIME_PERIOD")


def test_matrix_constant_is_defaulted_keyword_only_sequence(tmp_path: Path) -> None:
    workbook = _profile_workbook(tmp_path)
    modules = generate_inverted(workbook, _profile_bindings())
    assert "EvalContext" not in modules["api.py"]
    assert "ctx" not in modules["api.py"]
    assert "PROFILE_TABLE" in modules["data.py"]
    assert "(10.0, 11.0, 20.0, 21.0)" in modules["data.py"]
    pkg = load_package(modules, tmp_path, name="a8_kw")
    params = inspect.signature(pkg.compute_output_cell).parameters
    assert all(p.kind is inspect.Parameter.KEYWORD_ONLY for p in params.values())
    assert required_param_names(pkg.compute_output_cell) == ()
    assert "profile_table" in all_param_names(pkg.compute_output_cell)
    assert "ctx" not in all_param_names(pkg.compute_output_cell)
    assert _measure(pkg.compute_output_cell()) == pytest.approx(10.0)
    assert _measure(
        pkg.compute_output_cell(profile_table=(99.0, 11.0, 20.0, 21.0))
    ) == pytest.approx(99.0)


def test_matrix_cell_ref_matches_formula_evaluator(tmp_path: Path) -> None:
    workbook = _profile_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _profile_bindings()), tmp_path, name="a8_num")
    graph = create_dependency_graph(workbook, ["Outputs!A1"], load_values=True)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])
    assert _measure(pkg.compute_output_cell()) == pytest.approx(expected["Outputs!A1"])


def test_unknown_layout_still_fail_closed() -> None:
    with pytest.raises(InvertedTreeExportError, match="unsupported layout"):
        _layout_of({"id": "x", "layout": "grid"})
