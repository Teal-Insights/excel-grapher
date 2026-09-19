"""Bound label cells in Excel write-back.

Overlay copies `row_label`, `column_header`, `kind: cell`, and attribute
source cells missing from the view. Formula labels are written only when
every static dependency is already in the view or is itself a bound label.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import fastpyxl
import pytest

from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.series_bindings import bound_label_addresses, validate_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.fixtures.series_bindings.grouped_matrix_helpers import (
    grouped_matrix_bindings_document,
    write_grouped_matrix_workbook,
)


def _cell_value(path: Path, sheet: str, coord: str) -> object:
    wb = fastpyxl.load_workbook(path)
    try:
        return wb[sheet][coord].value
    finally:
        wb.close()


def _write_labelled_source(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = "indicator"
    ws["B1"] = 2024
    ws["C1"] = 2025
    ws["A2"] = "GDP"
    ws["B2"] = 1.1
    ws["C2"] = 1.2
    ws["Z99"] = "stray"
    wb.save(path)
    wb.close()


def _series_bindings(
    *,
    data_range: str = "Sheet1!B2:C2",
    dimensions: list[dict[str, Any]] | None = None,
    attributes: list[dict[str, Any]] | None = None,
    extra_concepts: list[dict[str, Any]] | None = None,
) -> WorkbookSeriesBindings:
    concepts = [
        {"id": "TIME_PERIOD", "dtype": "int"},
        {"id": "INDICATOR", "dtype": "string"},
        {"id": "OBS_VALUE", "dtype": "number"},
    ]
    if extra_concepts:
        concepts.extend(extra_concepts)
    if dimensions is None:
        dimensions = [
            {
                "concept": "INDICATOR",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
            },
            {
                "concept": "TIME_PERIOD",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            },
        ]
    structure: dict[str, Any] = {
        "measure": {
            "concept": "OBS_VALUE",
            "dtype": "float",
            "bind": {"kind": "data_cell", "read": "float"},
        },
        "dimensions": dimensions,
    }
    if attributes is not None:
        structure["attributes"] = attributes
    key = [str(dim["concept"]) for dim in dimensions]
    return validate_bindings_document(
        {
            "schema_version": "1.17.0",
            "concept_scheme": {"id": "writeback_labels", "concepts": concepts},
            "series": [
                {
                    "id": "gdp",
                    "sheet": "Sheet1",
                    "data_range": data_range,
                    "layout": "series",
                    "input": {},
                    "structure": structure,
                    "key": key,
                }
            ],
        }
    )


def _labelled_bindings() -> WorkbookSeriesBindings:
    return _series_bindings()


def _write_formula_header_source(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2024
    ws["B1"] = "=A1+1"
    ws["C1"] = 2026
    ws["B2"] = 1.5
    ws["C2"] = 1.6
    ws["Z99"] = "stray"
    wb.save(path)
    wb.close()


def _formula_header_bindings(*, data_range: str = "Sheet1!B2") -> WorkbookSeriesBindings:
    return _series_bindings(
        data_range=data_range,
        dimensions=[
            {
                "concept": "TIME_PERIOD",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            }
        ],
    )


def _extract_targets(source: Path, bindings: WorkbookSeriesBindings):
    return create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
    )


def test_write_workbook_without_bindings_omits_unreferenced_labels(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = _extract_targets(source, bindings)
    assert "Sheet1!A2" not in graph
    assert "Sheet1!B1" not in graph

    write_workbook(graph, dest)

    assert _cell_value(dest, "Sheet1", "B2") == 1.1
    assert _cell_value(dest, "Sheet1", "A2") is None
    assert _cell_value(dest, "Sheet1", "B1") is None
    assert _cell_value(dest, "Sheet1", "Z99") is None


def test_write_workbook_includes_bound_labels_from_bindings(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = _extract_targets(source, bindings)
    original_keys = set(graph)

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert set(graph) == original_keys
    assert _cell_value(dest, "Sheet1", "B2") == 1.1
    assert _cell_value(dest, "Sheet1", "C2") == 1.2
    assert _cell_value(dest, "Sheet1", "A2") == "GDP"
    assert _cell_value(dest, "Sheet1", "B1") == 2024
    assert _cell_value(dest, "Sheet1", "C1") == 2025
    assert _cell_value(dest, "Sheet1", "A1") is None
    assert _cell_value(dest, "Sheet1", "Z99") is None


def test_write_workbook_refuses_formula_label_when_deps_are_missing(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    _write_formula_header_source(source)
    bindings = _formula_header_bindings()
    graph = _extract_targets(source, bindings)
    assert "Sheet1!B1" not in graph
    assert "Sheet1!A1" not in graph

    with pytest.raises(ValueError, match=r"Sheet1!B1.*Sheet1!A1"):
        write_workbook(
            graph,
            tmp_path / "out.xlsx",
            series_bindings=bindings,
            bindings_workbook=source,
        )
    assert not (tmp_path / "out.xlsx").exists()


def test_write_workbook_writes_formula_label_when_deps_are_in_view(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_formula_header_source(source)
    bindings = _formula_header_bindings()
    graph = create_dependency_graph(
        source,
        [*all_series_targets(bindings, workbook=source), "Sheet1!A1"],
        load_values=True,
    )
    original_keys = set(graph)
    assert "Sheet1!B1" not in graph
    assert "Sheet1!A1" in graph

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert set(graph) == original_keys
    assert _cell_value(dest, "Sheet1", "B2") == 1.5
    assert _cell_value(dest, "Sheet1", "B1") == "=A1+1"
    assert _cell_value(dest, "Sheet1", "A1") == 2024
    assert _cell_value(dest, "Sheet1", "Z99") is None


def test_write_workbook_writes_formula_label_when_dep_is_bound_label(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2024
    ws["B1"] = "=C1"
    ws["C1"] = 2026
    ws["B2"] = 1.5
    ws["C2"] = 1.6
    ws["Z99"] = "stray"
    wb.save(source)
    wb.close()
    bindings = _formula_header_bindings(data_range="Sheet1!B2:C2")
    graph = _extract_targets(source, bindings)
    assert "Sheet1!B1" not in graph
    assert "Sheet1!C1" not in graph

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Sheet1", "B2") == 1.5
    assert _cell_value(dest, "Sheet1", "C2") == 1.6
    assert _cell_value(dest, "Sheet1", "B1") == "=C1"
    assert _cell_value(dest, "Sheet1", "C1") == 2026
    assert _cell_value(dest, "Sheet1", "A1") is None
    assert _cell_value(dest, "Sheet1", "Z99") is None


def test_write_workbook_refuses_offset_formula_label(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2024
    ws["B1"] = "=OFFSET(A1,0,0)"
    ws["B2"] = 1.5
    wb.save(source)
    wb.close()
    bindings = _formula_header_bindings()
    graph = create_dependency_graph(
        source,
        [*all_series_targets(bindings, workbook=source), "Sheet1!A1"],
        load_values=True,
    )

    with pytest.raises(ValueError, match=r"OFFSET|INDIRECT"):
        write_workbook(
            graph,
            tmp_path / "out.xlsx",
            series_bindings=bindings,
            bindings_workbook=source,
        )
    assert not (tmp_path / "out.xlsx").exists()


def test_write_workbook_bound_labels_require_workbook(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = _extract_targets(source, bindings)
    with pytest.raises(ValueError, match="bindings_workbook"):
        write_workbook(graph, tmp_path / "out.xlsx", series_bindings=bindings)


def test_write_workbook_include_bound_labels_false_omits_labels(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = _extract_targets(source, bindings)

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
        include_bound_labels=False,
    )

    assert _cell_value(dest, "Sheet1", "B2") == 1.1
    assert _cell_value(dest, "Sheet1", "A2") is None
    assert _cell_value(dest, "Sheet1", "B1") is None


def test_write_workbook_include_bound_labels_false_does_not_require_workbook(
    tmp_path: Path,
) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = _extract_targets(source, bindings)

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        include_bound_labels=False,
    )

    assert _cell_value(dest, "Sheet1", "B2") == 1.1
    assert _cell_value(dest, "Sheet1", "A2") is None


def test_write_workbook_graph_value_wins_over_workbook_label(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = create_dependency_graph(
        source,
        [*all_series_targets(bindings, workbook=source), "Sheet1!A2"],
        load_values=True,
    )
    graph.set_node_value("Sheet1!A2", "mutated")

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Sheet1", "A2") == "mutated"
    assert _cell_value(dest, "Sheet1", "B1") == 2024


def test_write_workbook_require_shared_formulas_with_bound_labels(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    _write_labelled_source(source)
    bindings = _labelled_bindings()
    graph = create_dependency_graph(
        source,
        all_series_targets(bindings, workbook=source),
        load_values=True,
        warm_formula_shapes=True,
    )

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
        shared_formulas="require",
    )

    assert _cell_value(dest, "Sheet1", "A2") == "GDP"
    assert _cell_value(dest, "Sheet1", "B2") == 1.1


def test_write_workbook_includes_kind_cell_and_attribute_binds(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = "region"
    ws["B1"] = 2024
    ws["A2"] = "GDP"
    ws["B2"] = 1.1
    ws["Z1"] = "note-cell"
    ws["Z2"] = "note-attr"
    wb.save(source)
    wb.close()
    bindings = _series_bindings(
        data_range="Sheet1!B2",
        dimensions=[
            {
                "concept": "INDICATOR",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
            },
            {
                "concept": "TIME_PERIOD",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            },
            {
                "concept": "REGION",
                "role": "key",
                "scope": "series",
                "bind": {"kind": "cell", "address": "Sheet1!Z1", "read": "string"},
            },
        ],
        attributes=[
            {
                "concept": "NOTE",
                "role": "attribute",
                "bind": {"kind": "cell", "address": "Sheet1!Z2", "read": "string"},
            }
        ],
        extra_concepts=[
            {"id": "REGION", "dtype": "string"},
            {"id": "NOTE", "dtype": "string"},
        ],
    )
    graph = _extract_targets(source, bindings)
    addresses = bound_label_addresses(bindings, workbook=source, graph=graph)
    assert "Sheet1!Z1" in addresses
    assert "Sheet1!Z2" in addresses

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Sheet1", "Z1") == "note-cell"
    assert _cell_value(dest, "Sheet1", "Z2") == "note-attr"
    assert _cell_value(dest, "Sheet1", "A2") == "GDP"
    assert _cell_value(dest, "Sheet1", "B1") == 2024


def test_write_workbook_projection_still_includes_bound_labels(tmp_path: Path) -> None:
    from excel_grapher.exporter import IdentityTransitCompression

    source = tmp_path / "src.xlsx"
    dest = tmp_path / "out.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 2024
    ws["B1"] = 2025
    ws["A2"] = 10
    ws["B2"] = "=A2"
    ws["B3"] = "=B2+1"
    wb.save(source)
    wb.close()

    bindings = _series_bindings(
        data_range="Sheet1!B3",
        dimensions=[
            {
                "concept": "TIME_PERIOD",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            }
        ],
    )
    graph = create_dependency_graph(
        source,
        ["Sheet1!B3"],
        capture_dependency_provenance=True,
        load_values=True,
    )
    projection = IdentityTransitCompression().project(graph)
    assert "Sheet1!B2" not in projection

    write_workbook(
        projection,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Sheet1", "B3") == "=A2+1"
    assert _cell_value(dest, "Sheet1", "B2") is None
    assert _cell_value(dest, "Sheet1", "A2") == 10
    assert _cell_value(dest, "Sheet1", "B1") == 2025
    assert _cell_value(dest, "Sheet1", "A1") is None


def test_write_workbook_projection_refuses_formula_label_to_collapsed_cell(
    tmp_path: Path,
) -> None:
    from excel_grapher.exporter import IdentityTransitCompression

    source = tmp_path / "src.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A2"] = 10
    ws["B1"] = "=B2"
    ws["B2"] = "=A2"
    ws["B3"] = "=B2+1"
    wb.save(source)
    wb.close()

    bindings = _formula_header_bindings(data_range="Sheet1!B3")
    graph = create_dependency_graph(
        source,
        ["Sheet1!B3"],
        capture_dependency_provenance=True,
        load_values=True,
    )
    projection = IdentityTransitCompression().project(graph)
    assert "Sheet1!B2" not in projection

    with pytest.raises(ValueError, match=r"Sheet1!B1.*Sheet1!B2"):
        write_workbook(
            projection,
            tmp_path / "out.xlsx",
            series_bindings=bindings,
            bindings_workbook=source,
        )
    assert not (tmp_path / "out.xlsx").exists()


def test_bound_label_addresses_uses_actual_fill_sources(tmp_path: Path) -> None:
    source = tmp_path / "grouped_inputs.xlsx"
    write_grouped_matrix_workbook(source)
    bindings = validate_bindings_document(grouped_matrix_bindings_document())
    graph = _extract_targets(source, bindings)

    addresses = bound_label_addresses(bindings, workbook=source, graph=graph)

    assert addresses == {
        "Inputs!A2",
        "Inputs!A3",
        "Inputs!A4",
        "Inputs!A6",
        "Inputs!A7",
        "Inputs!A8",
        "Inputs!C1",
        "Inputs!D1",
    }


def test_write_workbook_grouped_fill_labels(tmp_path: Path) -> None:
    source = tmp_path / "grouped_inputs.xlsx"
    dest = tmp_path / "out.xlsx"
    write_grouped_matrix_workbook(source)
    bindings = validate_bindings_document(grouped_matrix_bindings_document())
    graph = _extract_targets(source, bindings)

    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    assert _cell_value(dest, "Inputs", "A2") == "Paris"
    assert _cell_value(dest, "Inputs", "A3") == "Revenue"
    assert _cell_value(dest, "Inputs", "A4") == "Primary expenditure"
    assert _cell_value(dest, "Inputs", "A6") == "Moderate"
    assert _cell_value(dest, "Inputs", "A7") == "Revenue"
    assert _cell_value(dest, "Inputs", "A8") == "Primary expenditure"
    assert _cell_value(dest, "Inputs", "C1") == 2024
    assert _cell_value(dest, "Inputs", "D1") == 2025
    assert _cell_value(dest, "Inputs", "A5") is None
    assert _cell_value(dest, "Inputs", "C3") == 1.1
