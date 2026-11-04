"""Formula-header inputs on one axis may be sibling windows of one timeline.

The widest covering input is the labeller when one exists. A window that
sits outside that input is labelled from its own header cells.
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


def _input(series_id: str, data_range: str) -> dict:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction="input",
        header_row=2,
    )


def _output(series_id: str, data_range: str, *, scalar: bool = False) -> dict:
    if scalar:
        return series_entry(series_id, data_range, direction="output")
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction="output",
        header_row=2,
    )


def _assert_matches_evaluator(model: object, series: BoundSeries, graph: object) -> None:
    expected = FormulaEvaluator(graph).evaluate(list(series.cells))
    got = getattr(model, series.series_id)
    if series.single_valued:
        assert got == pytest.approx(expected[series.cells[0]])
        return
    assert tuple(got.domain.axes[0].keys) == tuple(point["TIME_PERIOD"] for point in series.domain)
    for (_coord, value), cell in zip(got.items(), series.cells, strict=True):
        assert value == pytest.approx(expected[cell])


def test_widest_formula_header_covers_sibling_windows(tmp_path: Path) -> None:
    """`early` and `late` do not nest, and both sit inside `span`."""
    path = tmp_path / "years.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A1": 2024,
                "A2": "=A1",
                "B2": "=A2+1",
                "C2": "=B2+1",
                "D2": "=C2+1",
                "A3": 1,
                "B3": 2,
                "C3": 3,
                "D3": 4,
                "A4": 10,
                "B4": 20,
                "C5": 30,
                "D5": 40,
                "E3": "=A3+B3+C3+D3",
                "E4": "=A4+B4",
                "A6": "=A4",
                "B6": "=B4",
                "C6": "=C5",
                "D6": "=D5",
            }
        },
    )
    _plant_years(path, {"A2": 2024, "B2": 2025, "C2": 2026, "D2": 2027})
    document = bindings_document(
        _input("span", "Inputs!A3:D3"),
        _input("early", "Inputs!A4:B4"),
        _input("late", "Inputs!C5:D5"),
        _output("total", "Inputs!E3", scalar=True),
        _output("early_total", "Inputs!E4", scalar=True),
        _output("window", "Inputs!A6:D6"),
    )
    pkg = load_package(
        generate_inverted(path, document),
        tmp_path,
        name="sibling_headers",
    )
    catalog, _deps, graph = inverted_graph_parts(path, document)
    model = pkg.Model.from_workbook(path)
    for series_id in ("total", "early_total", "window"):
        _assert_matches_evaluator(model, catalog.get(series_id), graph)
    assert tuple(model.span.domain.axes[0].keys) == (2024, 2025, 2026, 2027)
    assert tuple(model.early.domain.axes[0].keys) == (2024, 2025)
    assert tuple(model.late.domain.axes[0].keys) == (2026, 2027)
    assert dict(model.cells("early")) == {(2024,): "Inputs!A4", (2025,): "Inputs!B4"}
    assert dict(model.cells("late")) == {(2026,): "Inputs!C5", (2027,): "Inputs!D5"}

    case = tmp_path / "case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A1": 2030,
                "A2": "=A1",
                "B2": "=A2+1",
                "C2": "=B2+1",
                "D2": "=C2+1",
                "A3": 1,
                "B3": 2,
                "C3": 3,
                "D3": 4,
                "A4": 10,
                "B4": 20,
                "C5": 30,
                "D5": 40,
                "E3": "=A3+B3+C3+D3",
                "E4": "=A4+B4",
                "A6": "=A4",
                "B6": "=B4",
                "C6": "=C5",
                "D6": "=D5",
            }
        },
    )
    _plant_years(case, {"A2": 2030, "B2": 2031, "C2": 2032, "D2": 2033})
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(case, document)
    shifted = pkg.Model.from_workbook(case)
    for series_id in ("total", "early_total", "window"):
        _assert_matches_evaluator(shifted, shifted_catalog.get(series_id), shifted_graph)
    assert tuple(shifted.early.domain.axes[0].keys) == (2030, 2031)
    assert tuple(shifted.late.domain.axes[0].keys) == (2032, 2033)
    assert tuple(shifted.window.domain.axes[0].keys) == (2030, 2031, 2032, 2033)
    assert dict(shifted.cells("early")) == {(2030,): "Inputs!A4", (2031,): "Inputs!B4"}
    assert dict(shifted.cells("window")) == {
        (2030,): "Inputs!A6",
        (2031,): "Inputs!B6",
        (2032,): "Inputs!C6",
        (2033,): "Inputs!D6",
    }


def test_formula_header_windows_outside_the_widest_use_their_own_cells(tmp_path: Path) -> None:
    """History extends before the widest projection window, so neither nests."""
    path = tmp_path / "disjoint.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A1": 2011,
                "A2": "=A1",
                "B2": "=A2+1",
                "C2": "=B2+1",
                "E2": "=C2+11",
                "F2": "=E2+1",
                "G2": "=F2+1",
                "H2": "=G2+1",
                "A3": 1,
                "B3": 2,
                "C3": 3,
                "E3": 4,
                "F3": 5,
                "G3": 6,
                "H3": 7,
                "A4": "=A3+1",
                "B4": "=B3+1",
                "C4": "=C3+1",
                "E4": "=E3+1",
                "F4": "=E3+1",
                "G4": "=F3+1",
                "H4": "=G3+1",
                "A5": "=C3+H3",
            }
        },
    )
    _plant_years(
        path,
        {"A2": 2011, "B2": 2012, "C2": 2013, "E2": 2024, "F2": 2025, "G2": 2026, "H2": 2027},
    )
    document = bindings_document(
        series_entry("anchor", "Inputs!A1", direction="input", dtype="int"),
        _input("history", "Inputs!A3:C3"),
        _input("projection", "Inputs!E3:H3"),
        _output("history_out", "Inputs!A4:C4"),
        _output("projection_out", "Inputs!E4:H4"),
        _output("mixed", "Inputs!A5", scalar=True),
    )
    pkg = load_package(
        generate_inverted(path, document),
        tmp_path,
        name="disjoint_headers",
    )
    catalog, _deps, graph = inverted_graph_parts(path, document)
    model = pkg.Model.from_workbook(path)
    for series_id in ("history_out", "projection_out", "mixed"):
        _assert_matches_evaluator(model, catalog.get(series_id), graph)
    assert dict(model.cells("history")) == {
        (2011,): "Inputs!A3",
        (2012,): "Inputs!B3",
        (2013,): "Inputs!C3",
    }
    assert dict(model.cells("projection")) == {
        (2024,): "Inputs!E3",
        (2025,): "Inputs!F3",
        (2026,): "Inputs!G3",
        (2027,): "Inputs!H3",
    }

    case = tmp_path / "disjoint_case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A1": 2015,
                "A2": "=A1",
                "B2": "=A2+1",
                "C2": "=B2+1",
                "E2": "=C2+11",
                "F2": "=E2+1",
                "G2": "=F2+1",
                "H2": "=G2+1",
                "A3": 1,
                "B3": 2,
                "C3": 3,
                "E3": 4,
                "F3": 5,
                "G3": 6,
                "H3": 7,
                "A4": "=A3+1",
                "B4": "=B3+1",
                "C4": "=C3+1",
                "E4": "=E3+1",
                "F4": "=E3+1",
                "G4": "=F3+1",
                "H4": "=G3+1",
                "A5": "=C3+H3",
            }
        },
    )
    _plant_years(
        case,
        {"A2": 2015, "B2": 2016, "C2": 2017, "E2": 2028, "F2": 2029, "G2": 2030, "H2": 2031},
    )
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(case, document)
    shifted = pkg.Model.from_workbook(case)
    for series_id in ("history_out", "projection_out", "mixed"):
        _assert_matches_evaluator(shifted, shifted_catalog.get(series_id), shifted_graph)
    assert tuple(shifted.history.domain.axes[0].keys) == (2015, 2016, 2017)
    assert tuple(shifted.projection.domain.axes[0].keys) == (2028, 2029, 2030, 2031)
    assert dict(shifted.cells("history")) == {
        (2015,): "Inputs!A3",
        (2016,): "Inputs!B3",
        (2017,): "Inputs!C3",
    }
    assert shifted.anchor == 2015
