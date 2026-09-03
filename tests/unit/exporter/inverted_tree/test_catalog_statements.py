"""Catalog statements: per-cell key domains and auto shape-partition."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.resolve import resolve_key_domain
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a8_matrix import (
    _profile_bindings,
    _profile_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a12_formula_shape import (
    _a12_bindings,
    _a12_workbook,
)


def test_time_period_domain_resolves_header_values(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    catalog = build_catalog(validate_bindings_document(_a12_bindings()), workbook=workbook)
    series = catalog.get("path")
    assert [point["TIME_PERIOD"] for point in series.domain] == [2009, 2010, 2011]
    assert len(series.domain) == len(series.cells)


def test_matrix_domain_is_country_by_year_key_points(tmp_path: Path) -> None:
    workbook = _profile_workbook(tmp_path)
    catalog = build_catalog(validate_bindings_document(_profile_bindings()), workbook=workbook)
    series = catalog.get("profile_table")
    assert [point.as_mapping() for point in series.domain] == [
        {"COUNTRY": "France", "TIME_PERIOD": 2020},
        {"COUNTRY": "France", "TIME_PERIOD": 2021},
        {"COUNTRY": "Kenya", "TIME_PERIOD": 2020},
        {"COUNTRY": "Kenya", "TIME_PERIOD": 2021},
    ]


def test_mixed_formulas_partition_into_one_statement_per_shape_run(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    bindings = validate_bindings_document(_a12_bindings())
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("path")
    assert [stmt.shape_key for stmt in series.statements] == [
        series.statements[0].shape_key,
        series.statements[1].shape_key,
        series.statements[2].shape_key,
    ]
    assert len({stmt.shape_key for stmt in series.statements}) == 3
    assert [(stmt.start, stmt.stop) for stmt in series.statements] == [(0, 1), (1, 2), (2, 3)]
    assert [stmt.statement_id for stmt in series.statements] == [
        "path__0",
        "path__1",
        "path__2",
    ]
    assert series.statements[0].domain[0]["TIME_PERIOD"] == 2009


def test_uniform_formula_series_is_one_statement(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "uniform.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "A2": "=1",
                "B2": "=1",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "path",
            "Engine!A2:B2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("path")
    assert len(series.statements) == 1
    assert series.statements[0].statement_id == "path"
    assert series.statements[0].shape_key is not None
    assert series.statements[0].start == 0
    assert series.statements[0].stop == 2


def test_typo_bind_kind_fails_closed_instead_of_empty_domain(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    document = _a12_bindings()
    document["series"][0]["structure"]["dimensions"][0]["bind"]["kind"] = "colum_header"
    entry = document["series"][0]
    with pytest.raises(ValueError, match="key field"):
        resolve_key_domain(workbook, entry, ("Engine!A2", "Engine!B2", "Engine!C2"))
    with pytest.raises(InvertedTreeExportError, match="key field"):
        build_catalog(document, workbook=workbook)
