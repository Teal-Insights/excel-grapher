"""A scalar-layout series with one cell per scenario sheet is a keyed series."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    write_workbook,
)


def _sheets_workbook(tmp_path: Path) -> Path:
    sheets: dict[str, dict[str, object]] = {}
    for name, shock in (("S1", 1.5), ("S2", 2.5)):
        cells: dict[str, object] = {"A1": shock, "B1": 2024, "C1": 2025, "D1": 2026}
        for offset, column in enumerate("BCD"):
            cells[f"{column}2"] = float(offset + 1)
            cells[f"{column}3"] = f"={column}2*$A$1"
        sheets[name] = cells
    return write_workbook(tmp_path / "sheets.xlsx", sheets)


def _scenario_dimension() -> dict[str, Any]:
    return {
        "id": "SCENARIO",
        "concept": "SCENARIO",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "sheet_name", "values": {"S1": "s1", "S2": "s2"}},
    }


def _period_dimension() -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
    }


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _sheets_bindings() -> dict[str, Any]:
    shock = {
        "id": "shock",
        "sheet": ["S1", "S2"],
        "data_range": ["S1!A1", "S2!A1"],
        "layout": "scalar",
        "input": {"setter": {"name": "set_shock"}},
        "structure": {"measure": _measure(), "dimensions": [_scenario_dimension()]},
        "key": ["SCENARIO"],
    }
    flow = {
        "id": "flow",
        "sheet": ["S1", "S2"],
        "data_range": ["S1!B2:D2", "S2!B2:D2"],
        "layout": "matrix",
        "input": {"setter": {"name": "set_flow"}},
        "structure": {
            "measure": _measure(),
            "dimensions": [_scenario_dimension(), _period_dimension()],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    shocked = {
        **flow,
        "id": "shocked",
        "data_range": ["S1!B3:D3", "S2!B3:D3"],
        "output": {"compute": {"name": "compute_shocked"}},
    }
    del shocked["input"]
    return bindings_document(shock, flow, shocked, schema_version="1.15.0")


def test_scalar_layout_series_keyed_by_sheet_is_a_keyed_series(tmp_path: Path) -> None:
    modules = generate_inverted(_sheets_workbook(tmp_path), _sheets_bindings())
    data, internals, api = modules["data.py"], modules["internals.py"], modules["api.py"]
    assert "class Shock(Series[float | str | None]):" in data
    assert "SHOCK_DEFAULT = Shock(SHOCK_DOMAIN, (1.5, 2.5))" in data
    assert "xl_mul(flow[scenario, time_period], shock[scenario])" in internals
    assert "shock: data.Shock" in api
    pkg = assert_package_matches_evaluator(
        _sheets_workbook(tmp_path), _sheets_bindings(), tmp_path, "keyed_scalars"
    )
    shocked = pkg.compute_shocked(flow=pkg.data.FLOW_DEFAULT, shock=pkg.data.SHOCK_DEFAULT)
    assert shocked["s2", 2026] == 3.0 * 2.5
