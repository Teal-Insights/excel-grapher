"""Same-year reads of two scenarios are keyed accesses, not a catalog lag (#733).

A host member that reads `gdp` at this-scenario and baseline for the same
`TIME_PERIOD` is `gdp[Baseline, t] / gdp[Stress, t]`. Catalog order concatenates
scenario blocks, so those members are not adjacent; packing them next to each
other (`hi - lo == 1`) is still not a 1-period lag.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a10_other_series_lag import (
    _non_lag_bindings,
    _non_lag_workbook,
)

_TIME_DIM = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 8, "read": "int"},
}


def _scenario_dim(values: dict[str, str]) -> dict[str, Any]:
    return {
        "id": "SCENARIO",
        "concept": "SCENARIO",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "sheet_name", "values": values},
    }


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _mcve_sheets() -> dict[str, dict[str, object]]:
    return {
        "Stress": {
            "E8": 2024,
            "F8": 2025,
            "E46": 50.0,
            "F46": 55.0,
            "E19": "=Baseline!O19*Baseline!O48/Stress!E46",
            "F19": "=Baseline!P19*Baseline!P48/Stress!F46",
        },
        "Baseline": {
            "O8": 2024,
            "P8": 2025,
            "O19": 10.0,
            "P19": 11.0,
            "O48": 200.0,
            "P48": 220.0,
        },
        "Results": {
            "E8": 2024,
            "F8": 2025,
            "E19": "=Stress!E19",
            "F19": "=Stress!F19",
        },
    }


def _gdp_entry(data_range: str | list[str]) -> dict[str, Any]:
    sheets = ["Stress", "Baseline"]
    if isinstance(data_range, list):
        sheets = [item.split("!", 1)[0] for item in data_range]
    entry: dict[str, Any] = {
        "id": "gdp",
        "sheet": sheets,
        "data_range": data_range,
        "layout": "series",
        "input": {"setter": {"name": "set_gdp", "record_contract": "records", "strict": True}},
        "structure": {
            "measure": _measure(),
            "dimensions": [
                _scenario_dim({"Stress": "B1", "Baseline": "Baseline"}),
                _TIME_DIM,
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    return entry


def _mcve_bindings(*, gdp_range: str | list[str] | None = None) -> dict[str, Any]:
    baseline = series_entry(
        "baseline_exports",
        "Baseline!O19:P19",
        layout="series",
        direction="input",
        header_row=8,
    )
    result = series_entry(
        "result",
        "Results!E19:F19",
        layout="series",
        direction="output",
        header_row=8,
    )
    internal: dict[str, Any] = {
        "id": "scaled_exports",
        "sheet": "Stress",
        "data_range": "Stress!E19:F19",
        "layout": "series",
        "internal": {},
        "structure": {"measure": _measure(), "dimensions": [_TIME_DIM]},
        "key": ["TIME_PERIOD"],
    }
    return bindings_document(
        baseline,
        _gdp_entry(gdp_range or ["Stress!E46:F46", "Baseline!O48:P48"]),
        result,
        internal,
        schema_version="1.14.0",
    )


def _multi_scenario_sheets() -> dict[str, dict[str, object]]:
    sheets = _mcve_sheets()
    sheets["Shock"] = {
        "E8": 2024,
        "F8": 2025,
        "E46": 60.0,
        "F46": 66.0,
        "E19": "=Baseline!O19*Baseline!O48/Shock!E46",
        "F19": "=Baseline!P19*Baseline!P48/Shock!F46",
    }
    return sheets


def _multi_scenario_bindings() -> dict[str, Any]:
    baseline = series_entry(
        "baseline_exports",
        "Baseline!O19:P19",
        layout="series",
        direction="input",
        header_row=8,
    )
    gdp: dict[str, Any] = {
        "id": "gdp",
        "sheet": ["Stress", "Shock", "Baseline"],
        "data_range": ["Stress!E46:F46", "Shock!E46:F46", "Baseline!O48:P48"],
        "layout": "series",
        "input": {"setter": {"name": "set_gdp"}},
        "structure": {
            "measure": _measure(),
            "dimensions": [
                _scenario_dim({"Stress": "B1", "Shock": "B2", "Baseline": "Baseline"}),
                _TIME_DIM,
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    output: dict[str, Any] = {
        "id": "scaled_exports",
        "sheet": ["Stress", "Shock"],
        "data_range": ["Stress!E19:F19", "Shock!E19:F19"],
        "layout": "series",
        "output": {"compute": {"name": "compute_scaled_exports"}},
        "structure": {
            "measure": _measure(),
            "dimensions": [
                _scenario_dim({"Stress": "B1", "Shock": "B2"}),
                _TIME_DIM,
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    return bindings_document(baseline, gdp, output, schema_version="1.14.0")


def test_keyed_dual_read_is_not_a_lag(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a29.xlsx", _mcve_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings())
    scaled = deps["scaled_exports"]
    assert "gdp" in scaled.param_ids
    assert "gdp" in scaled.keyed_ids
    assert "gdp" not in scaled.lagged_ids
    assert "gdp" not in scaled.aligned_ids
    assert catalog.get("gdp").cells == (
        "Stress!E46",
        "Stress!F46",
        "Baseline!O48",
        "Baseline!P48",
    )


def test_keyed_dual_read_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a29_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a29_eval")
    cells = ["Results!E19", "Results!F19"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((40.0, 44.0))


def test_adjacent_scenario_pack_is_not_a_lag(tmp_path: Path) -> None:
    """Baseline packed next to stress (`hi - lo == 1`) is still two keyed reads."""
    workbook = write_workbook(tmp_path / "a29_adj.xlsx", _mcve_sheets())
    document = _mcve_bindings(
        gdp_range=["Stress!E46", "Baseline!O48", "Stress!F46", "Baseline!P48"]
    )
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    scaled = deps["scaled_exports"]
    assert catalog.get("gdp").cells == (
        "Stress!E46",
        "Baseline!O48",
        "Stress!F46",
        "Baseline!P48",
    )
    assert "gdp" in scaled.keyed_ids
    assert "gdp" not in scaled.lagged_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a29_adj")
    cells = ["Results!E19", "Results!F19"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((40.0, 44.0))
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "gdp[i + 1]" not in internals


def test_multi_scenario_host_keyed_dual_read_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a29_host.xlsx", _multi_scenario_sheets())
    document = _multi_scenario_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    scaled = deps["scaled_exports"]
    assert "gdp" in scaled.keyed_ids
    assert "gdp" not in scaled.lagged_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a29_host")
    cells = ["Stress!E19", "Stress!F19", "Shock!E19", "Shock!F19"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_scaled_exports(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((40.0, 44.0, 10.0 * 200.0 / 60.0, 11.0 * 220.0 / 66.0))


def test_unclassifiable_two_positions_name_cells_and_keys(tmp_path: Path) -> None:
    workbook = _non_lag_workbook(tmp_path)
    with pytest.raises(InvertedTreeExportError, match=r"Engine!A2.*TIME_PERIOD=2009") as exc:
        generate_inverted(workbook, _non_lag_bindings())
    message = str(exc.value)
    assert "Engine!C2" in message
    assert "TIME_PERIOD=2011" in message
    assert "direction" in message
    assert "debt" in message
    assert "(2, 0)" not in message
