"""Aligned plus pinned-year reads of one TIME_PERIOD series are keyed (#735).

A fade `baseline[t] - (baseline[2026] - shock0) * …` is two accesses, not a
catalog-index lag. Adjacent years (`hi - lo == 1`) can be this shape; treating
that pair as `saw_lag` would emit `baseline[i + 1]` instead of `baseline[2026]`.

#733's same-year outer-key helper does not apply: both reads share the series
and scenario and differ only on `TIME_PERIOD`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
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

_TIME_DIM = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 8, "read": "int"},
}


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _mcve_sheets() -> dict[str, dict[str, object]]:
    return {
        "Baseline": {
            "E8": 2025,
            "F8": 2026,
            "G8": 2027,
            "H8": 2028,
            "E2": 10.0,
            "F2": 11.0,
            "G2": 12.0,
            "H2": 13.0,
        },
        "Stress": {
            "E8": 2025,
            "F8": 2026,
            "G8": 2027,
            "H8": 2028,
            "A1": 9.0,
            "A2": 6,
            "E19": "=Baseline!E2",
            "F19": "=Baseline!F2-(Baseline!$F$2-Stress!$A$1)*($A$2-1)/$A$2",
            "G19": "=Baseline!G2-(Baseline!$F$2-Stress!$A$1)*($A$2-2)/$A$2",
            "H19": "=Baseline!H2-(Baseline!$F$2-Stress!$A$1)*($A$2-3)/$A$2",
        },
        "Results": {
            "E8": 2025,
            "F8": 2026,
            "G8": 2027,
            "H8": 2028,
            "E19": "=Stress!E19",
            "F19": "=Stress!F19",
            "G19": "=Stress!G19",
            "H19": "=Stress!H19",
        },
    }


def _series(sid: str, sheets: str, data_range: str, extra: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": sid,
        "sheet": sheets,
        "data_range": data_range,
        "layout": "series",
        **extra,
        "structure": {"measure": _measure(), "dimensions": [_TIME_DIM]},
        "key": ["TIME_PERIOD"],
    }


def _mcve_bindings() -> dict[str, Any]:
    return bindings_document(
        _series(
            "baseline",
            "Baseline",
            "Baseline!E2:H2",
            {
                "input": {
                    "setter": {"name": "set_baseline", "record_contract": "records", "strict": True}
                }
            },
        ),
        series_entry("shock0", "Stress!A1", layout="scalar", direction="input"),
        series_entry("horizon", "Stress!A2", layout="scalar", direction="input", dtype="int"),
        _series(
            "result",
            "Results",
            "Results!E19:H19",
            {"output": {"compute": {"name": "compute_result", "record_contract": "records"}}},
        ),
        _series("faded", "Stress", "Stress!E19:H19", {"internal": {}}),
        schema_version="1.14.0",
    )


def _adjacent_cousin_sheets() -> dict[str, dict[str, object]]:
    """Host is only 2027: aligned G2 plus pinned $F$2 (`hi - lo == 1`)."""
    return {
        "Baseline": {
            "E8": 2025,
            "F8": 2026,
            "G8": 2027,
            "H8": 2028,
            "E2": 10.0,
            "F2": 11.0,
            "G2": 12.0,
            "H2": 13.0,
        },
        "Stress": {
            "E8": 2027,
            "A1": 9.0,
            "A2": 6,
            "E19": "=Baseline!G2-(Baseline!$F$2-Stress!$A$1)*($A$2-2)/$A$2",
        },
        "Results": {
            "E8": 2027,
            "E19": "=Stress!E19",
        },
    }


def _adjacent_cousin_bindings() -> dict[str, Any]:
    return bindings_document(
        _series(
            "baseline",
            "Baseline",
            "Baseline!E2:H2",
            {"input": {"setter": {"name": "set_baseline"}}},
        ),
        series_entry("shock0", "Stress!A1", layout="scalar", direction="input"),
        series_entry("horizon", "Stress!A2", layout="scalar", direction="input", dtype="int"),
        _series(
            "result",
            "Results",
            "Results!E19",
            {"output": {"compute": {"name": "compute_result"}}},
        ),
        _series("faded", "Stress", "Stress!E19", {"internal": {}}),
        schema_version="1.14.0",
    )


def test_pinned_year_dual_read_is_not_a_lag(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a30.xlsx", _mcve_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings())
    faded = deps["faded"]
    assert "baseline" in faded.param_ids
    assert "baseline" in faded.keyed_ids
    assert "baseline" not in faded.lagged_ids
    assert "baseline" not in faded.aligned_ids
    assert catalog.get("baseline").cells == (
        "Baseline!E2",
        "Baseline!F2",
        "Baseline!G2",
        "Baseline!H2",
    )


def test_pinned_year_dual_read_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a30_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a30_eval")
    cells = ["Results!E19", "Results!F19", "Results!G19", "Results!H19"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    # E19 copies 10. F19: 11 - (11 - 9) * 5 / 6 = 11 - 10/6.
    # G19: 12 - (11 - 9) * 4 / 6 = 12 - 8/6. H19: 13 - (11 - 9) * 3 / 6 = 12.
    assert got == pytest.approx((10.0, 11.0 - 10.0 / 6.0, 12.0 - 8.0 / 6.0, 12.0))
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "baseline[i + 1]" not in internals
    assert "baseline[i - 1]" not in internals


def test_adjacent_pinned_year_is_not_a_lag(tmp_path: Path) -> None:
    """2026 and 2027 (`hi - lo == 1`) is still aligned + pin, not `saw_lag`."""
    workbook = write_workbook(tmp_path / "a30_adj.xlsx", _adjacent_cousin_sheets())
    document = _adjacent_cousin_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    faded = deps["faded"]
    assert catalog.get("baseline").cells == (
        "Baseline!E2",
        "Baseline!F2",
        "Baseline!G2",
        "Baseline!H2",
    )
    assert "baseline" in faded.keyed_ids
    assert "baseline" not in faded.lagged_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a30_adj")
    cells = ["Results!E19"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx((expected["Results!E19"],))
    assert got == pytest.approx((12.0 - 8.0 / 6.0,))
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "baseline[i + 1]" not in internals


def _scenario_dim() -> dict[str, Any]:
    return {
        "id": "SCENARIO",
        "concept": "SCENARIO",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "sheet_name", "values": {"Stress": "B1", "Baseline": "Baseline"}},
    }


def _cross_shape_sheets() -> dict[str, dict[str, object]]:
    return {
        "Stress": {
            "E8": 2025,
            "F8": 2026,
            "E46": 50.0,
            "F46": 55.0,
            "E19": "=Stress!E46/Baseline!$F$48",
            "F19": "=Stress!F46/Baseline!$F$48",
        },
        "Baseline": {
            "E8": 2025,
            "F8": 2026,
            "E48": 200.0,
            "F48": 220.0,
        },
        "Results": {
            "E8": 2025,
            "F8": 2026,
            "E19": "=Stress!E19",
            "F19": "=Stress!F19",
        },
    }


def _cross_shape_bindings() -> dict[str, Any]:
    gdp: dict[str, Any] = {
        "id": "gdp",
        "sheet": ["Stress", "Baseline"],
        "data_range": ["Stress!E46:F46", "Baseline!E48:F48"],
        "layout": "series",
        "input": {"setter": {"name": "set_gdp"}},
        "structure": {
            "measure": _measure(),
            "dimensions": [_scenario_dim(), _TIME_DIM],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    return bindings_document(
        gdp,
        _series(
            "result",
            "Results",
            "Results!E19:F19",
            {"output": {"compute": {"name": "compute_result"}}},
        ),
        _series("scaled", "Stress", "Stress!E19:F19", {"internal": {}}),
        schema_version="1.14.0",
    )


def test_aligned_scenario_plus_pinned_year_is_keyed(tmp_path: Path) -> None:
    """This-scenario `[s, t]` plus pinned `baseline[2026]` is two keyed reads."""
    workbook = write_workbook(tmp_path / "a30_cross.xlsx", _cross_shape_sheets())
    document = _cross_shape_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    scaled = deps["scaled"]
    assert "gdp" in scaled.keyed_ids
    assert "gdp" not in scaled.lagged_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a30_cross")
    cells = ["Results!E19", "Results!F19"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((50.0 / 220.0, 55.0 / 220.0))
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "gdp[i + 1]" not in internals
    assert catalog.get("gdp").cells == (
        "Stress!E46",
        "Stress!F46",
        "Baseline!E48",
        "Baseline!F48",
    )
