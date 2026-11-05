"""A formula-header labeller contributes the axis it labels, at any rank.

The widest covering input can be a matrix whose first axis is some other
key. Position indexes, template sources, and provenance bind that labelled
axis by name.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_model_from_workbook import _plant_formula_cache


def _plant_years(path: Path, years: dict[str, int]) -> None:
    for coordinate, year in years.items():
        _plant_formula_cache(path, coordinate, str(year))


def _service_series() -> dict:
    return {
        "id": "service",
        "sheet": "Inputs",
        "data_range": "Inputs!A10:D11",
        "layout": "series",
        "input": {},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "INSTRUMENT",
                    "concept": "INSTRUMENT",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "value_map",
                        "values": {"IMF": 10, "IDA": 11},
                        "read": "string",
                    },
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 2, "read": "int"},
                },
            ],
        },
        "key": ["INSTRUMENT", "TIME_PERIOD"],
    }


def _bindings() -> dict:
    document = bindings_document(
        _service_series(),
        series_entry(
            "path",
            "Inputs!A5:D5",
            layout="series",
            direction="output",
            header_row=2,
        ),
        series_entry(
            "window",
            "Inputs!A6:B6",
            layout="series",
            direction="output",
            header_row=2,
        ),
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "string"})
    return document


def _sheet(origin: int) -> dict[str, object]:
    sheet: dict[str, object] = {
        "A1": origin,
        "A2": "=A1",
        "B2": "=A2+1",
        "C2": "=B2+1",
        "D2": "=C2+1",
        "A5": "=A10+A11",
        "B5": "=B10+B11",
        "C5": "=C10",
        "D5": "=D10",
        "A6": "=A10+A11",
        "B6": "=B10+B11",
    }
    for column, values in zip("ABCD", ((1, 2), (3, 4), (5, 6), (7, 8)), strict=True):
        sheet[f"{column}10"] = values[0]
        sheet[f"{column}11"] = values[1]
    return sheet


def _write(path: Path, origin: int) -> None:
    write_workbook(path, {"Inputs": _sheet(origin)})
    _plant_years(
        path,
        {
            "A2": origin,
            "B2": origin + 1,
            "C2": origin + 2,
            "D2": origin + 3,
        },
    )


def _assert_matches_evaluator(model: object, series: BoundSeries, graph: object) -> None:
    expected = FormulaEvaluator(graph).evaluate(list(series.cells))
    got = getattr(model, series.series_id)
    assert tuple(axis.name for axis in got.domain.axes) == ("TIME_PERIOD",)
    assert tuple(got.domain.axes[0].keys) == tuple(point["TIME_PERIOD"] for point in series.domain)
    for (_coord, value), cell in zip(got.items(), series.cells, strict=True):
        assert value == pytest.approx(expected[cell])


def test_matrix_formula_header_labeller_labels_by_axis_name(tmp_path: Path) -> None:
    """`service` is keyed by INSTRUMENT, then TIME_PERIOD, and labels both outputs."""
    path = tmp_path / "years.xlsx"
    _write(path, 2024)
    document = _bindings()
    modules = generate_inverted(path, document)
    assert ".domain.axis('TIME_PERIOD')" in modules["internals.py"]
    assert "source=(2024, 2025, 2026, 2027)" in modules["data.py"]
    pkg = load_package(modules, tmp_path, name="matrix_labeller")
    catalog, _deps, graph = inverted_graph_parts(path, document)
    service = catalog.get("service")
    assert tuple(axis.name for axis in service.tensor_domain.axes) == (
        "INSTRUMENT",
        "TIME_PERIOD",
    )
    model = pkg.Model.from_workbook(path)
    for series_id in ("path", "window"):
        _assert_matches_evaluator(model, catalog.get(series_id), graph)
    assert tuple(model.service.domain.axes[1].keys) == (2024, 2025, 2026, 2027)
    assert dict(model.cells("window")) == {(2024,): "Inputs!A6", (2025,): "Inputs!B6"}
    assert model.cells("service")[("IMF", 2024)] == "Inputs!A10"
    assert model.cells("service")[("IDA", 2027)] == "Inputs!D11"

    case = tmp_path / "case.xlsx"
    _write(case, 2030)
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(case, document)
    shifted = pkg.Model.from_workbook(case)
    for series_id in ("path", "window"):
        _assert_matches_evaluator(shifted, shifted_catalog.get(series_id), shifted_graph)
    assert tuple(shifted.service.domain.axes[1].keys) == (2030, 2031, 2032, 2033)
    assert tuple(shifted.window.domain.axes[0].keys) == (2030, 2031)
    assert dict(shifted.cells("window")) == {(2030,): "Inputs!A6", (2031,): "Inputs!B6"}
    assert shifted.cells("service")[("IDA", 2033)] == "Inputs!D11"


def _two_axis_series(series_id: str, data_range: str, *, direction: str) -> dict:
    entry: dict = {
        "id": series_id,
        "sheet": "Inputs",
        "data_range": data_range,
        "layout": "series",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "INSTRUMENT",
                    "concept": "INSTRUMENT",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "int"},
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["INSTRUMENT", "TIME_PERIOD"],
    }
    if direction == "input":
        entry["input"] = {}
    else:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    return entry


def _two_axis_bindings() -> dict:
    document = bindings_document(
        _two_axis_series("grid", "Inputs!B2:C3", direction="input"),
        _two_axis_series("copy", "Inputs!B5:C6", direction="output"),
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "int"})
    return document


def _two_axis_sheet(origin: int) -> dict[str, object]:
    return {
        "Z1": origin,
        "A1": 1,
        "B1": "=Z1",
        "C1": "=B1+1",
        "A2": "=A1",
        "A3": "=A2+1",
        "B2": 10,
        "C2": 20,
        "B3": 30,
        "C3": 40,
        "A5": "=A2",
        "A6": "=A3",
        "B5": "=B2",
        "C5": "=C2+1",
        "B6": "=B3",
        "C6": "=C3+1",
    }


def _write_two_axis(path: Path, origin: int) -> None:
    write_workbook(path, {"Inputs": _two_axis_sheet(origin)})
    _plant_years(
        path,
        {
            "B1": origin,
            "C1": origin + 1,
            "A2": 1,
            "A3": 2,
            "A5": 1,
            "A6": 2,
        },
    )


def test_one_matrix_labeller_supplies_each_formula_header_axis(tmp_path: Path) -> None:
    """Row and column formula headers on one input each bind their own axis."""
    path = tmp_path / "grid.xlsx"
    _write_two_axis(path, 2024)
    document = _two_axis_bindings()
    modules = generate_inverted(path, document)
    internals = modules["internals.py"]
    assert "_ax_instrument = grid.domain.axis('INSTRUMENT')" in internals
    assert "_ax_time_period = grid.domain.axis('TIME_PERIOD')" in internals
    pkg = load_package(modules, tmp_path, name="two_axis_labeller")
    catalog, _deps, graph = inverted_graph_parts(path, document)
    model = pkg.Model.from_workbook(path)
    series = catalog.get("copy")
    expected = FormulaEvaluator(graph).evaluate(list(series.cells))
    got = model.copy
    for (coord, value), point, cell in zip(got.items(), series.domain, series.cells, strict=True):
        assert coord == tuple(point[field] for field in series.key_fields)
        assert value == pytest.approx(expected[cell])
    assert model.cells("grid")[(1, 2024)] == "Inputs!B2"
    assert model.cells("copy")[(2, 2025)] == "Inputs!C6"

    case = tmp_path / "case.xlsx"
    _write_two_axis(case, 2030)
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(case, document)
    shifted = pkg.Model.from_workbook(case)
    shifted_series = shifted_catalog.get("copy")
    shifted_expected = FormulaEvaluator(shifted_graph).evaluate(list(shifted_series.cells))
    for (coord, value), point, cell in zip(
        shifted.copy.items(), shifted_series.domain, shifted_series.cells, strict=True
    ):
        assert coord == tuple(point[field] for field in shifted_series.key_fields)
        assert value == pytest.approx(shifted_expected[cell])
    assert shifted.cells("copy")[(2, 2031)] == "Inputs!C6"
