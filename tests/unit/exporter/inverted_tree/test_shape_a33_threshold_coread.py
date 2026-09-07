"""Keyed dual-reads stay keyed when one slot is a constant scenario (#741).

A host member that compares a path scenario to a cap is
`1 if paths[s, t] > paths[Threshold, t] else 0`. Host rows bind `SCENARIO` as
`Baseline breach` / `Shock breach`; the producer binds the same concept as
`Baseline` / `MX shock Standard&Tailored` / `Threshold`.

`_host_follow_key_maps` used to require exactly one producer value per member.
The extra `Threshold` slot dropped `SCENARIO` from the follow map, both slots
became literals, and `_is_keyed_multi_read` fail-closed because the path
literals disagree (`Baseline` vs `MX shock…`).

`Threshold` is the intersection of every member's reads, so it stays a
literal. The remaining path value is a function of the host row, and emit
remaps it through the follow map.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)

_TIME_DIM = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
}

_PATH_SCENARIO = {
    "id": "SCENARIO",
    "concept": "SCENARIO",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "value_map",
        "values": {
            "Baseline": 5,
            "MX shock Standard&Tailored": 7,
            "Threshold": 9,
        },
        "read": "string",
    },
}

_BREACH_SCENARIO = {
    "id": "SCENARIO",
    "concept": "SCENARIO",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "value_map",
        "values": {
            "Baseline breach": 11,
            "Shock breach": 12,
        },
        "read": "string",
    },
}


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _unwrap(value: object) -> object:
    """Return the sole member of a 1-tuple compute result."""
    if isinstance(value, tuple) and len(value) == 1:
        return value[0]
    return value


def _mcve_sheets() -> dict[str, dict[str, object]]:
    return {
        "Chart": {
            "D1": 2024,
            "E1": 2025,
            "D5": 5.0,
            "E5": 6.0,
            "D7": 8.0,
            "E7": 9.0,
            "D9": 4.0,
            "E9": 7.0,
            "D11": "=IF(Chart!D5>Chart!D9,1,0)",
            "E11": "=IF(Chart!E5>Chart!E9,1,0)",
            "D12": "=IF(Chart!D7>Chart!D9,1,0)",
            "E12": "=IF(Chart!E7>Chart!E9,1,0)",
        },
        "Outputs": {
            "A1": "=Chart!D11",
            "A2": "=Chart!E11",
            "A3": "=Chart!D12",
            "A4": "=Chart!E12",
        },
    }


def _paths_entry() -> dict[str, Any]:
    return {
        "id": "paths",
        "sheet": "Chart",
        "data_range": "Chart!D5:E9",
        "layout": "series",
        "exclude_rows": ["6", "8"],
        "input": {
            "setter": {
                "name": "set_paths",
                "record_contract": "records",
                "strict": True,
            }
        },
        "structure": {
            "measure": _measure(),
            "dimensions": [_PATH_SCENARIO, _TIME_DIM],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }


def _breach_entry() -> dict[str, Any]:
    return {
        "id": "breach",
        "sheet": "Chart",
        "data_range": ["Chart!D11:E11", "Chart!D12:E12"],
        "layout": "series",
        "internal": {},
        "structure": {
            "measure": _measure(),
            "dimensions": [_BREACH_SCENARIO, _TIME_DIM],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }


def _mcve_bindings() -> dict[str, Any]:
    return bindings_document(
        _paths_entry(),
        series_entry(
            "result_baseline_year",
            "Outputs!A1",
            layout="scalar",
            direction="output",
        ),
        series_entry(
            "result_shock_year",
            "Outputs!A3",
            layout="scalar",
            direction="output",
        ),
        _breach_entry(),
        schema_version="1.14.0",
    )


def test_threshold_coread_is_keyed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a33.xlsx", _mcve_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings())
    host = deps["breach"]
    assert "paths" in host.param_ids
    assert "paths" in host.keyed_ids
    assert "paths" not in host.lagged_ids
    assert "paths" not in host.aligned_ids
    assert catalog.get("paths").cells == (
        "Chart!D5",
        "Chart!E5",
        "Chart!D7",
        "Chart!E7",
        "Chart!D9",
        "Chart!E9",
    )


def test_threshold_coread_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a33_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "paths" in deps["breach"].keyed_ids
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "Threshold" in internals
    assert "Baseline" in internals
    assert "MX shock Standard&Tailored" in internals
    pkg = load_package(modules, tmp_path, name="a33_eval")
    cells = ["Outputs!A1", "Outputs!A2", "Outputs!A3", "Outputs!A4"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = pkg.internals.breach(kwargs["paths"])
    assert got == pytest.approx((1.0, 0.0, 1.0, 1.0))
    assert (
        _unwrap(call_compute(pkg, "result_baseline_year", kwargs)),
        _unwrap(call_compute(pkg, "result_shock_year", kwargs)),
    ) == pytest.approx((expected["Outputs!A1"], expected["Outputs!A3"]))
