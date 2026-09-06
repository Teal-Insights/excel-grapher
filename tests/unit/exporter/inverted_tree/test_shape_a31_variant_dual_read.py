"""`$` pins Excel axes, not sheet-bound keys (#737).

A host that reads two `VARIANT`s of a multi-scenario matrix at `$D$2` / `$F$2`
is `stats[s, indicator, mean]` and `stats[s, indicator, stdev]`. The `$` freezes
the column (variant) and row (indicator); the sheet *is* the host scenario, so
`SCENARIO` stays `host` even though the formula spells `Stress!$D$2`.

#735 treated every `$` pin as a literal for every key. That disagrees across
`sheet_name` hosts (`SCENARIO='shock'` vs `'combo'`) and fail-closes. A
one-scenario host that hits the same catalog slots still emits via
`static_catalog` (affine slope 0). The multi-scenario host is the miss.

The same rule applies on the other axis: two `row_label` indicators pinned
with `$` still follow the host sheet.
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
    "bind": {"kind": "column_header", "header_row": 8, "read": "int"},
}

_TIME_ROW_DIM = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "row_label", "label_column": "A", "read": "int"},
}

_VARIANT_DIM = {
    "id": "VARIANT",
    "concept": "VARIANT",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "value_map",
        "values": {"Historical average": "D", "Standard deviation": "F"},
    },
}

_INDICATOR_DIM = {
    "id": "INDICATOR",
    "concept": "INDICATOR",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "row_label",
        "label_column": "B",
        "read": "string",
        "normalize": "strip",
    },
}

_INDICATOR_ROW_DIM = {
    "id": "INDICATOR",
    "concept": "INDICATOR",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "row_label",
        "label_column": "A",
        "read": "string",
        "normalize": "strip",
    },
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


def _unwrap(value: object) -> object:
    """Return the sole member of a 1-tuple compute result."""
    if isinstance(value, tuple) and len(value) == 1:
        return value[0]
    return value


def _document(*series: dict[str, Any]) -> dict[str, Any]:
    doc = bindings_document(*series, schema_version="1.14.0")
    known = {item["id"] for item in doc["concept_scheme"]["concepts"]}
    for extra in ("INDICATOR", "VARIANT"):
        if extra not in known:
            doc["concept_scheme"]["concepts"].append({"id": extra, "dtype": "string"})
    return doc


def _mcve_sheets() -> dict[str, dict[str, object]]:
    sheets: dict[str, dict[str, object]] = {
        "Params": {"C1": 1.0},
        "Outputs": {"A1": "=Stress!E19", "A2": "=Combo!E19"},
    }
    for title in ("Stress", "Combo"):
        sheets[title] = {
            "E8": 2025,
            "B2": "FDI to GDP",
            "D2": 7.0,
            "F2": 2.0,
            "E19": f"=-({title}!$D$2-Params!$C$1*{title}!$F$2)",
        }
    return sheets


def _stats_entry(*, sheets: list[str], data_range: list[str]) -> dict[str, Any]:
    return {
        "id": "stats",
        "sheet": sheets,
        "data_range": data_range,
        "layout": "matrix",
        "exclude_columns": ["E"],
        "input": {"setter": {"name": "set_stats", "record_contract": "records", "strict": True}},
        "structure": {
            "measure": _measure(),
            "dimensions": [
                _scenario_dim(
                    {name: name.lower() if name != "Stress" else "shock" for name in sheets}
                ),
                _INDICATOR_DIM,
                _VARIANT_DIM,
            ],
        },
        "key": ["SCENARIO", "INDICATOR", "VARIANT"],
    }


def _mcve_bindings() -> dict[str, Any]:
    fdi = {
        "id": "fdi",
        "sheet": ["Stress", "Combo"],
        "data_range": ["Stress!E19", "Combo!E19"],
        "layout": "series",
        "internal": {},
        "structure": {
            "measure": _measure(),
            "dimensions": [
                _scenario_dim({"Stress": "shock", "Combo": "combo"}),
                _TIME_DIM,
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    return _document(
        _stats_entry(sheets=["Stress", "Combo"], data_range=["Stress!D2:F2", "Combo!D2:F2"]),
        series_entry("shock_size", "Params!C1", layout="scalar", direction="input"),
        series_entry("result_shock", "Outputs!A1", layout="scalar", direction="output"),
        series_entry("result_combo", "Outputs!A2", layout="scalar", direction="output"),
        fdi,
    )


def _two_year_sheets() -> dict[str, dict[str, object]]:
    sheets = _mcve_sheets()
    for title in ("Stress", "Combo"):
        sheets[title]["F8"] = 2026
        sheets[title]["F19"] = f"=-({title}!$D$2-Params!$C$1*{title}!$F$2)"
    sheets["Outputs"]["B1"] = "=Stress!F19"
    sheets["Outputs"]["B2"] = "=Combo!F19"
    return sheets


def _two_year_bindings() -> dict[str, Any]:
    doc = _mcve_bindings()
    for series in doc["series"]:
        if series["id"] == "fdi":
            series["data_range"] = ["Stress!E19:F19", "Combo!E19:F19"]
    return doc


def _row_indicator_sheets() -> dict[str, dict[str, object]]:
    sheets: dict[str, dict[str, object]] = {
        "Params": {"C1": 1.0},
        "Outputs": {"A1": "=Stress!B19", "A2": "=Combo!B19"},
    }
    for title in ("Stress", "Combo"):
        sheets[title] = {
            "A19": 2025,
            "A2": "Historical average",
            "B2": 7.0,
            "A4": "Standard deviation",
            "B4": 2.0,
            "B19": f"=-({title}!$B$2-Params!$C$1*{title}!$B$4)",
        }
    return sheets


def _row_indicator_bindings() -> dict[str, Any]:
    stats = {
        "id": "stats",
        "sheet": ["Stress", "Combo"],
        "data_range": ["Stress!B2", "Stress!B4", "Combo!B2", "Combo!B4"],
        "layout": "matrix",
        "input": {"setter": {"name": "set_stats"}},
        "structure": {
            "measure": _measure(),
            "dimensions": [
                _scenario_dim({"Stress": "shock", "Combo": "combo"}),
                _INDICATOR_ROW_DIM,
            ],
        },
        "key": ["SCENARIO", "INDICATOR"],
    }
    host = {
        "id": "fdi",
        "sheet": ["Stress", "Combo"],
        "data_range": ["Stress!B19", "Combo!B19"],
        "layout": "series",
        "internal": {},
        "structure": {
            "measure": _measure(),
            "dimensions": [
                _scenario_dim({"Stress": "shock", "Combo": "combo"}),
                _TIME_ROW_DIM,
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    return _document(
        stats,
        series_entry("shock_size", "Params!C1", layout="scalar", direction="input"),
        series_entry("result_shock", "Outputs!A1", layout="scalar", direction="output"),
        series_entry("result_combo", "Outputs!A2", layout="scalar", direction="output"),
        host,
    )


def _cross_sheet_sheets() -> dict[str, dict[str, object]]:
    sheets = _mcve_sheets()
    sheets["Baseline"] = {
        "E8": 2025,
        "B2": "FDI to GDP",
        "D2": 10.0,
        "F2": 4.0,
    }
    sheets["Stress"]["E19"] = "=Stress!$D$2/Baseline!$D$2"
    sheets["Combo"]["E19"] = "=Combo!$D$2/Baseline!$D$2"
    return sheets


def _cross_sheet_bindings() -> dict[str, Any]:
    doc = _mcve_bindings()
    for series in doc["series"]:
        if series["id"] == "stats":
            series["sheet"] = ["Stress", "Combo", "Baseline"]
            series["data_range"] = ["Stress!D2:F2", "Combo!D2:F2", "Baseline!D2:F2"]
            series["structure"]["dimensions"][0] = _scenario_dim(
                {"Stress": "shock", "Combo": "combo", "Baseline": "baseline"}
            )
    return doc


def test_variant_dual_read_is_keyed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a31.xlsx", _mcve_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings())
    fdi = deps["fdi"]
    assert "stats" in fdi.param_ids
    assert "stats" in fdi.keyed_ids
    assert "stats" not in fdi.lagged_ids
    assert "stats" not in fdi.aligned_ids
    assert catalog.get("stats").cells == (
        "Stress!D2",
        "Stress!F2",
        "Combo!D2",
        "Combo!F2",
    )


def test_variant_dual_read_emits_host_scenario_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a31_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "(('shock', 'combo')[i], 'FDI to GDP', 'Historical average')" in internals
    assert "(('shock', 'combo')[i], 'FDI to GDP', 'Standard deviation')" in internals
    pkg = load_package(modules, tmp_path, name="a31_eval")
    cells = ["Outputs!A1", "Outputs!A2"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = (
        _unwrap(call_compute(pkg, "result_shock", kwargs)),
        _unwrap(call_compute(pkg, "result_combo", kwargs)),
    )
    assert got == pytest.approx((expected["Outputs!A1"], expected["Outputs!A2"]))
    assert got == pytest.approx((-5.0, -5.0))


def test_two_year_host_same_pins_is_keyed(tmp_path: Path) -> None:
    """Each year rereads `$D$2`/`$F$2`; catalog slots still differ by sheet."""
    workbook = write_workbook(tmp_path / "a31_years.xlsx", _two_year_sheets())
    catalog, deps, graph = inverted_graph_parts(workbook, _two_year_bindings())
    fdi = deps["fdi"]
    assert "stats" in fdi.keyed_ids
    assert "stats" not in fdi.lagged_ids
    document = _two_year_bindings()
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a31_years")
    cells = ["Stress!E19", "Stress!F19", "Combo!E19", "Combo!F19"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = pkg.internals.fdi(kwargs["stats"], kwargs["shock_size"])
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "stats[i + 1]" not in internals
    assert catalog.get("stats").cells == (
        "Stress!D2",
        "Stress!F2",
        "Combo!D2",
        "Combo!F2",
    )


def test_row_pinned_indicator_dual_read_is_keyed(tmp_path: Path) -> None:
    """Two `$`-pinned row labels still bind `SCENARIO` as `host`."""
    workbook = write_workbook(tmp_path / "a31_rows.xlsx", _row_indicator_sheets())
    catalog, deps, graph = inverted_graph_parts(workbook, _row_indicator_bindings())
    fdi = deps["fdi"]
    assert "stats" in fdi.keyed_ids
    assert "stats" not in fdi.lagged_ids
    document = _row_indicator_bindings()
    modules = generate_inverted(workbook, document)
    assert "(('shock', 'combo')[i], 'Historical average')" in modules["internals.py"]
    assert "(('shock', 'combo')[i], 'Standard deviation')" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a31_rows")
    cells = ["Outputs!A1", "Outputs!A2"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = (
        _unwrap(call_compute(pkg, "result_shock", kwargs)),
        _unwrap(call_compute(pkg, "result_combo", kwargs)),
    )
    assert got == pytest.approx((expected["Outputs!A1"], expected["Outputs!A2"]))
    assert got == pytest.approx((-5.0, -5.0))
    assert catalog.get("stats").cells == (
        "Stress!B2",
        "Stress!B4",
        "Combo!B2",
        "Combo!B4",
    )


def test_cross_sheet_pin_keeps_other_scenario_literal(tmp_path: Path) -> None:
    """A `$` pin on another sheet still freezes that `sheet_name` key."""
    workbook = write_workbook(tmp_path / "a31_cross.xlsx", _cross_sheet_sheets())
    catalog, deps, graph = inverted_graph_parts(workbook, _cross_sheet_bindings())
    fdi = deps["fdi"]
    assert "stats" in fdi.keyed_ids
    assert "stats" not in fdi.lagged_ids
    document = _cross_sheet_bindings()
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "(('shock', 'combo')[i], 'FDI to GDP', 'Historical average')" in internals
    assert "stats[4]" in internals
    pkg = load_package(modules, tmp_path, name="a31_cross")
    cells = ["Outputs!A1", "Outputs!A2"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = (
        _unwrap(call_compute(pkg, "result_shock", kwargs)),
        _unwrap(call_compute(pkg, "result_combo", kwargs)),
    )
    assert got == pytest.approx((expected["Outputs!A1"], expected["Outputs!A2"]))
    assert got == pytest.approx((7.0 / 10.0, 7.0 / 10.0))
    assert catalog.get("stats").cells == (
        "Stress!D2",
        "Stress!F2",
        "Combo!D2",
        "Combo!F2",
        "Baseline!D2",
        "Baseline!F2",
    )
