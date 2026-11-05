"""A period-index TIME_PERIOD is not a key on a calendar-year labeller.

Both series share the axis name. The host coordinate is an index along the
columns (0, 1, 2, …) and the producer coordinate is the formula year in that
column. Export addresses the year by position on each series' own axis.
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


def _assert_matches_evaluator(model: object, series: BoundSeries, graph: object) -> None:
    expected = FormulaEvaluator(graph).evaluate(list(series.cells))
    got = getattr(model, series.series_id)
    assert tuple(got.domain.axes[0].keys) == tuple(point["TIME_PERIOD"] for point in series.domain)
    for (_coord, value), cell in zip(got.items(), series.cells, strict=True):
        assert value == pytest.approx(expected[cell])


def test_period_index_reads_calendar_year_series_by_column(tmp_path: Path) -> None:
    """Host keys 2 and 3 read amount years in columns A and B, two columns left."""
    path = tmp_path / "years.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A1": 2024,
                "B1": "=A1+1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "A2": 0,
                "B2": 1,
                "C2": 2,
                "D2": 3,
                "A3": 10,
                "B3": 20,
                "C3": 30,
                "D3": 40,
                "C4": "=A3",
                "D4": "=B3",
            }
        },
    )
    _plant_years(path, {"B1": 2025, "C1": 2026, "D1": 2027})
    document = bindings_document(
        series_entry(
            "amount",
            "Inputs!A3:D3",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "picked",
            "Inputs!C4:D4",
            layout="series",
            direction="output",
            header_row=2,
        ),
    )
    modules = generate_inverted(path, document)
    picked = modules["internals.py"].split("def picked", 1)[1].split("\ndef ", 1)[0]
    assert "2024" not in picked
    assert "2025" not in picked
    pkg = load_package(modules, tmp_path, name="period_index_pick")
    catalog, _deps, graph = inverted_graph_parts(path, document)
    model = pkg.Model.from_workbook(path)
    _assert_matches_evaluator(model, catalog.get("picked"), graph)
    assert tuple(model.amount.domain.axes[0].keys) == (2024, 2025, 2026, 2027)
    assert tuple(model.picked.domain.axes[0].keys) == (2, 3)
    assert model.picked[(2,)] == pytest.approx(10)
    assert model.picked[(3,)] == pytest.approx(20)

    case = tmp_path / "shifted.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A1": 2030,
                "B1": "=A1+1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "A2": 0,
                "B2": 1,
                "C2": 2,
                "D2": 3,
                "A3": 10,
                "B3": 20,
                "C3": 30,
                "D3": 40,
                "C4": "=A3",
                "D4": "=B3",
            }
        },
    )
    _plant_years(case, {"B1": 2031, "C1": 2032, "D1": 2033})
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(case, document)
    shifted = pkg.Model.from_workbook(case)
    _assert_matches_evaluator(shifted, shifted_catalog.get("picked"), shifted_graph)
    assert tuple(shifted.amount.domain.axes[0].keys) == (2030, 2031, 2032, 2033)
    assert tuple(shifted.picked.domain.axes[0].keys) == (2, 3)
    assert shifted.picked[(2,)] == pytest.approx(10)
    assert shifted.picked[(3,)] == pytest.approx(20)


def test_period_index_window_uses_the_shared_axis_constant(tmp_path: Path) -> None:
    """A host window that is not a prefix of the interned axis still aligns.

    Periods 0–3 and years 2024–2027 share one `TIME_PERIOD` constant. `picked`
    only occupies keys 2 and 3, and each cell reads one column to the left.
    """
    path = tmp_path / "window.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A1": 2024,
                "B1": "=A1+1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "A2": 0,
                "B2": 1,
                "C2": 2,
                "D2": 3,
                "A3": 10,
                "B3": 20,
                "C3": 30,
                "D3": 40,
                "A5": 1,
                "B5": 1,
                "C5": 1,
                "D5": 1,
                "C4": "=B3",
                "D4": "=C3",
            }
        },
    )
    _plant_years(path, {"B1": 2025, "C1": 2026, "D1": 2027})
    document = bindings_document(
        series_entry(
            "amount",
            "Inputs!A3:D3",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "index",
            "Inputs!A5:D5",
            layout="series",
            direction="input",
            header_row=2,
        ),
        series_entry(
            "picked",
            "Inputs!C4:D4",
            layout="series",
            direction="output",
            header_row=2,
        ),
    )
    modules = generate_inverted(path, document)
    axis = "Axis('TIME_PERIOD', (0, 1, 2, 3, 2024, 2025, 2026, 2027), int)"
    assert axis in modules["data.py"]
    picked = modules["internals.py"].split("def picked", 1)[1].split("\ndef ", 1)[0]
    assert "position(time_period) - 1" in picked
    pkg = load_package(modules, tmp_path, name="period_index_window")
    catalog, _deps, graph = inverted_graph_parts(path, document)
    model = pkg.Model.from_workbook(path)
    _assert_matches_evaluator(model, catalog.get("picked"), graph)
    assert model.picked[(2,)] == pytest.approx(20)
    assert model.picked[(3,)] == pytest.approx(30)


def test_period_index_sum_spans_calendar_years_in_the_same_columns(tmp_path: Path) -> None:
    """An expanding SUM and a running total use the year in the host column.

    Period headers are formulas (`0`, then `=+previous+1`) on an input, so
    the host axis is that labeller. The summed series is the calendar-year
    header on the row above. Shifting the year anchor moves year keys and
    leaves the period index and the column sums in place.
    """
    path = tmp_path / "sum.xlsx"

    def sheet(origin: int) -> dict[str, object]:
        return {
            "A1": origin,
            "A2": "=A1",
            "B2": "=A2+1",
            "C2": "=B2+1",
            "D2": "=C2+1",
            "A3": 0,
            "B3": "=A3+1",
            "C3": "=B3+1",
            "D3": "=C3+1",
            "A4": 10,
            "B4": 20,
            "C4": 30,
            "D4": 40,
            "A5": "=SUM($A4:A4)",
            "B5": "=SUM($A4:B4)",
            "C5": "=SUM($A4:C4)",
            "D5": "=SUM($A4:D4)",
            "A6": "=A4",
            "B6": "=A6+B4",
            "C6": "=B6+C4",
            "D6": "=C6+D4",
            "A7": 1,
            "B7": 1,
            "C7": 1,
            "D7": 1,
        }

    write_workbook(path, {"Inputs": sheet(2024)})
    _plant_years(
        path,
        {
            "A2": 2024,
            "B2": 2025,
            "C2": 2026,
            "D2": 2027,
            "B3": 1,
            "C3": 2,
            "D3": 3,
        },
    )
    document = bindings_document(
        series_entry(
            "amount",
            "Inputs!A4:D4",
            layout="series",
            direction="input",
            header_row=2,
        ),
        series_entry(
            "period",
            "Inputs!A7:D7",
            layout="series",
            direction="input",
            header_row=3,
        ),
        series_entry(
            "cumulative",
            "Inputs!A5:D5",
            layout="series",
            direction="output",
            header_row=3,
        ),
        series_entry(
            "running",
            "Inputs!A6:D6",
            layout="series",
            direction="output",
            header_row=3,
        ),
    )
    modules = generate_inverted(path, document)
    pkg = load_package(modules, tmp_path, name="period_index_sum")
    catalog, _deps, graph = inverted_graph_parts(path, document)
    model = pkg.Model.from_workbook(path)
    for series_id in ("cumulative", "running"):
        _assert_matches_evaluator(model, catalog.get(series_id), graph)
    assert [model.cumulative[(index,)] for index in range(4)] == pytest.approx([10, 30, 60, 100])
    assert [model.running[(index,)] for index in range(4)] == pytest.approx([10, 30, 60, 100])

    case = tmp_path / "sum_shifted.xlsx"
    write_workbook(case, {"Inputs": sheet(2030)})
    _plant_years(
        case,
        {
            "A2": 2030,
            "B2": 2031,
            "C2": 2032,
            "D2": 2033,
            "B3": 1,
            "C3": 2,
            "D3": 3,
        },
    )
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(case, document)
    shifted = pkg.Model.from_workbook(case)
    for series_id in ("cumulative", "running"):
        _assert_matches_evaluator(shifted, shifted_catalog.get(series_id), shifted_graph)
    assert tuple(shifted.amount.domain.axes[0].keys) == (2030, 2031, 2032, 2033)
    assert tuple(shifted.cumulative.domain.axes[0].keys) == (0, 1, 2, 3)
    assert [shifted.cumulative[(index,)] for index in range(4)] == pytest.approx([10, 30, 60, 100])
