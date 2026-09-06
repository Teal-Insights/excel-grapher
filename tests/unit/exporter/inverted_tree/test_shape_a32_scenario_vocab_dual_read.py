"""Dual instrument reads stay keyed when host and producer SCENARIO vocabularies differ (#739).

A host member that sums two `INSTRUMENT`s of one producer series at the matching
year is `amort[s, External, t] + amort[s, Domestic, t]`. Host sheets bind
`SCENARIO` as short codes (`B1` / `B2`); the producer binds the same concept as
full stress-test names. String equality then marks `SCENARIO` as a literal that
disagrees across host sheets, so `_is_keyed_multi_read` fail-closes.

The producer values still follow the host scenario: each host sheet reads one
producer block. Emit remaps host codes onto producer names and keeps the two
instrument lookups keyed.
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

_HOST_TIME = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 7, "read": "int"},
}

_PROD_TIME = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
}

_PROD_INSTRUMENT = {
    "id": "INSTRUMENT",
    "concept": "INSTRUMENT",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "value_map",
        "values": {
            "External PPG medium and long-term": [5, 11],
            "Domestic medium and long-term": [8, 14],
        },
        "read": "string",
    },
}

_PROD_SCENARIO = {
    "id": "SCENARIO",
    "concept": "SCENARIO",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "value_map",
        "values": {
            "Bounds Test 1: Real GDP Growth Shock": [5, 8],
            "Bounds Test 2: Primary Balance Shock": [11, 14],
        },
        "read": "string",
    },
}

_HOST_SCENARIO = {
    "id": "SCENARIO",
    "concept": "SCENARIO",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "sheet_name",
        "values": {"Host_B1": "B1", "Host_B2": "B2"},
    },
}

_ALIGNED_HOST_SCENARIO = {
    "id": "SCENARIO",
    "concept": "SCENARIO",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "sheet_name",
        "values": {
            "Host_B1": "Bounds Test 1: Real GDP Growth Shock",
            "Host_B2": "Bounds Test 2: Primary Balance Shock",
        },
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


def _document(*series: dict[str, Any]) -> dict[str, Any]:
    doc = bindings_document(*series, schema_version="1.14.0")
    known = {item["id"] for item in doc["concept_scheme"]["concepts"]}
    if "INSTRUMENT" not in known:
        doc["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "string"})
    return doc


def _mcve_sheets() -> dict[str, dict[str, object]]:
    return {
        "Producer": {
            "E1": 2025,
            "F1": 2026,
            "E5": 10.0,
            "F5": 11.0,
            "E8": 3.0,
            "F8": 4.0,
            "E11": 20.0,
            "F11": 21.0,
            "E14": 5.0,
            "F14": 6.0,
        },
        "Host_B1": {
            "E7": 2025,
            "F7": 2026,
            "E10": "=Producer!E5+Producer!E8",
            "F10": "=Producer!F5+Producer!F8",
        },
        "Host_B2": {
            "E7": 2025,
            "F7": 2026,
            "E10": "=Producer!E11+Producer!E14",
            "F10": "=Producer!F11+Producer!F14",
        },
        "Outputs": {
            "A1": "=Host_B1!E10",
            "A2": "=Host_B2!E10",
            "A3": "=Host_B1!F10",
            "A4": "=Host_B2!F10",
        },
    }


def _amortization_entry() -> dict[str, Any]:
    return {
        "id": "amortization",
        "sheet": "Producer",
        "data_range": "Producer!E5:F14",
        "layout": "series",
        "exclude_rows": ["6:7", "9:10", "12:13"],
        "input": {
            "setter": {
                "name": "set_amortization",
                "record_contract": "records",
                "strict": True,
            }
        },
        "structure": {
            "measure": _measure(),
            "dimensions": [_PROD_INSTRUMENT, _PROD_SCENARIO, _PROD_TIME],
        },
        "key": ["INSTRUMENT", "SCENARIO", "TIME_PERIOD"],
    }


def _host_entry(*, scenario: dict[str, Any] | None = None) -> dict[str, Any]:
    return {
        "id": "total_amortization",
        "sheet": ["Host_B1", "Host_B2"],
        "data_range": ["Host_B1!E10:F10", "Host_B2!E10:F10"],
        "layout": "series",
        "internal": {},
        "structure": {
            "measure": _measure(),
            "dimensions": [scenario or _HOST_SCENARIO, _HOST_TIME],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }


def _mcve_bindings(*, scenario: dict[str, Any] | None = None) -> dict[str, Any]:
    return _document(
        _amortization_entry(),
        series_entry("result_shock", "Outputs!A1", layout="scalar", direction="output"),
        series_entry("result_combo", "Outputs!A2", layout="scalar", direction="output"),
        _host_entry(scenario=scenario),
    )


def test_scenario_vocab_dual_read_is_keyed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a32.xlsx", _mcve_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings())
    host = deps["total_amortization"]
    assert "amortization" in host.param_ids
    assert "amortization" in host.keyed_ids
    assert "amortization" not in host.lagged_ids
    assert "amortization" not in host.aligned_ids
    assert catalog.get("amortization").cells == (
        "Producer!E5",
        "Producer!F5",
        "Producer!E8",
        "Producer!F8",
        "Producer!E11",
        "Producer!F11",
        "Producer!E14",
        "Producer!F14",
    )


def test_scenario_vocab_dual_read_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a32_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "amortization" in deps["total_amortization"].keyed_ids
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "External PPG medium and long-term" in internals
    assert "Domestic medium and long-term" in internals
    assert "Bounds Test 1: Real GDP Growth Shock" in internals
    assert "Bounds Test 2: Primary Balance Shock" in internals
    pkg = load_package(modules, tmp_path, name="a32_eval")
    cells = ["Outputs!A1", "Outputs!A2", "Host_B1!F10", "Host_B2!F10"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = pkg.internals.total_amortization(kwargs["amortization"])
    assert got == pytest.approx((13.0, 15.0, 25.0, 27.0))
    assert (
        _unwrap(call_compute(pkg, "result_shock", kwargs)),
        _unwrap(call_compute(pkg, "result_combo", kwargs)),
    ) == pytest.approx((expected["Outputs!A1"], expected["Outputs!A2"]))


def test_aligned_scenario_names_still_emit(tmp_path: Path) -> None:
    """Shared vocabulary is still `host`; the remap is a no-op."""
    workbook = write_workbook(tmp_path / "a32_aligned.xlsx", _mcve_sheets())
    document = _mcve_bindings(scenario=_ALIGNED_HOST_SCENARIO)
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "amortization" in deps["total_amortization"].keyed_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a32_aligned")
    got = pkg.internals.total_amortization(input_kwargs(catalog, graph)["amortization"])
    assert got == pytest.approx((13.0, 15.0, 25.0, 27.0))
