"""Complementary series shards across sheets (schema 1.14.0)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import xlsxwriter
import yaml

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    SeriesBindingsLoadError,
    SeriesBindingsSchemaError,
    expand_data_range,
    load_series_bindings,
    merge_series_binding_documents,
    resolve_series_binding,
    validate_bindings_document,
)
from excel_grapher.series_bindings.versions import (
    IMPLEMENTED_BIND_KINDS,
    SUPPORTED_SCHEMA_VERSIONS,
)


def _scenario_series(
    *,
    sheet: str,
    data_range: str,
    scenario: Any | None = None,
    bind: dict[str, Any] | None = None,
    direction: str = "internal",
) -> dict[str, Any]:
    if bind is None:
        bind = {"kind": "constant", "value": sheet if scenario is None else scenario}
    series: dict[str, Any] = {
        "id": "external_debt_nominal",
        "sheet": sheet,
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
                    "id": "SCENARIO",
                    "concept": "SCENARIO",
                    "role": "key",
                    "scope": "cell",
                    "bind": bind,
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 8, "read": "int"},
                },
            ],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }
    if direction == "internal":
        series["internal"] = {}
    elif direction == "constant":
        series["constant"] = {}
    elif direction == "output":
        series["output"] = {"compute": {"name": "compute_external_debt_nominal"}}
    else:
        series["input"] = {"setter": {"name": "set_external_debt_nominal"}}
    return series


def _scenario_doc(
    series: dict[str, Any] | list[dict[str, Any]],
    *,
    schema_version: str = "1.14.0",
) -> dict[str, Any]:
    series_list = series if isinstance(series, list) else [series]
    return {"schema_version": schema_version, "series": series_list}


def test_schema_version_1_14_0_supported() -> None:
    assert "1.14.0" in SUPPORTED_SCHEMA_VERSIONS
    assert "sheet_name" in IMPLEMENTED_BIND_KINDS


def test_schema_accepts_sheet_name_bind() -> None:
    series = _scenario_series(
        sheet="Baseline",
        data_range="Baseline!C12:D12",
        bind={"kind": "sheet_name"},
    )
    bindings = validate_bindings_document(_scenario_doc(series))
    assert bindings["series"][0]["structure"]["dimensions"][0]["bind"]["kind"] == "sheet_name"


def test_schema_accepts_sheet_name_values_map() -> None:
    series = _scenario_series(
        sheet="B1_GDP_ext",
        data_range="B1_GDP_ext!C12:D12",
        bind={"kind": "sheet_name", "values": {"B1_GDP_ext": "B1", "Baseline DSA": "Baseline"}},
    )
    validate_bindings_document(_scenario_doc(series))


def test_schema_accepts_list_data_range_and_sheets() -> None:
    series = _scenario_series(
        sheet="Baseline",
        data_range="Baseline!C12:D12",
        bind={"kind": "sheet_name"},
    )
    series["sheet"] = ["Baseline", "B1"]
    series["data_range"] = ["Baseline!C12:D12", "B1!C12:D12"]
    bindings = validate_bindings_document(_scenario_doc(series))
    assert bindings["series"][0]["data_range"] == ["Baseline!C12:D12", "B1!C12:D12"]
    assert bindings["series"][0]["sheet"] == ["Baseline", "B1"]


def test_schema_infers_sheet_list_from_list_data_range() -> None:
    series = _scenario_series(
        sheet="Baseline",
        data_range="Baseline!C12:D12",
        bind={"kind": "sheet_name"},
    )
    del series["sheet"]
    series["data_range"] = ["Baseline!C12:D12", "B1!C12:D12"]
    bindings = validate_bindings_document(_scenario_doc(series))
    assert bindings["series"][0]["sheet"] == ["Baseline", "B1"]


def test_schema_rejects_sheet_name_with_unknown_property() -> None:
    series = _scenario_series(
        sheet="Baseline",
        data_range="Baseline!C12:D12",
        bind={"kind": "sheet_name", "address": "Baseline!A1"},
    )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_scenario_doc(series))


def test_merge_complementary_shards_mcve(tmp_path: Path) -> None:
    """Issue 590 MCVE: same id, different sheet/range/SCENARIO constant."""
    a = _scenario_doc(
        _scenario_series(sheet="Baseline", data_range="Baseline!C12:X12", scenario="Baseline"),
        schema_version="1.13.0",
    )
    b = _scenario_doc(
        _scenario_series(sheet="B1", data_range="B1!C12:X12", scenario="B1"),
        schema_version="1.13.0",
    )
    shard_dir = tmp_path / "bindings"
    shard_dir.mkdir()
    (shard_dir / "a.bindings.yaml").write_text(yaml.safe_dump(a), encoding="utf-8")
    (shard_dir / "b.bindings.yaml").write_text(yaml.safe_dump(b), encoding="utf-8")

    bindings = load_series_bindings(shard_dir)
    assert len(bindings["series"]) == 1
    series = bindings["series"][0]
    assert series["id"] == "external_debt_nominal"
    assert series["sheet"] == ["Baseline", "B1"]
    assert series["data_range"] == ["Baseline!C12:X12", "B1!C12:X12"]
    assert series["key"] == ["SCENARIO", "TIME_PERIOD"]
    scenario_bind = series["structure"]["dimensions"][0]["bind"]
    assert scenario_bind["kind"] == "sheet_name"
    assert "values" not in scenario_bind


def test_merge_rejects_duplicate_id_within_one_shard() -> None:
    series = _scenario_series(sheet="Baseline", data_range="Baseline!C12:X12")
    doc = _scenario_doc([series, dict(series)])
    with pytest.raises(SeriesBindingsLoadError, match="Duplicate series id"):
        merge_series_binding_documents([doc])


def test_merge_maps_distinct_scenario_codes_via_sheet_name() -> None:
    left = _scenario_doc(
        _scenario_series(
            sheet="Baseline DSA",
            data_range="Baseline DSA!C12:X12",
            scenario="Baseline",
        )
    )
    right = _scenario_doc(
        _scenario_series(
            sheet="B1_GDP_ext",
            data_range="B1_GDP_ext!C12:X12",
            scenario="B1",
        )
    )
    merged = merge_series_binding_documents([left, right])
    series = merged["series"][0]
    assert series["sheet"] == ["Baseline DSA", "B1_GDP_ext"]
    bind = series["structure"]["dimensions"][0]["bind"]
    assert bind == {
        "kind": "sheet_name",
        "values": {"Baseline DSA": "Baseline", "B1_GDP_ext": "B1"},
    }


def test_merge_identical_sheet_name_structure_concatenates_ranges() -> None:
    bind = {"kind": "sheet_name"}
    left = _scenario_doc(
        _scenario_series(sheet="Baseline", data_range="Baseline!C12:X12", bind=bind)
    )
    right = _scenario_doc(_scenario_series(sheet="B1", data_range="B1!C12:X12", bind=bind))
    merged = merge_series_binding_documents([left, right])
    series = merged["series"][0]
    assert series["data_range"] == ["Baseline!C12:X12", "B1!C12:X12"]
    assert series["structure"]["dimensions"][0]["bind"] == {"kind": "sheet_name"}


def test_merge_keeps_shared_constant_when_values_agree() -> None:
    left = _scenario_doc(
        _scenario_series(sheet="Baseline", data_range="Baseline!C12:X12", scenario="LIC")
    )
    right = _scenario_doc(_scenario_series(sheet="B1", data_range="B1!C12:X12", scenario="LIC"))
    merged = merge_series_binding_documents([left, right])
    bind = merged["series"][0]["structure"]["dimensions"][0]["bind"]
    assert bind == {"kind": "constant", "value": "LIC"}


def test_merge_same_sheet_different_range_concatenates() -> None:
    left = _scenario_doc(
        _scenario_series(sheet="Inputs", data_range="Inputs!B2:C2", scenario="Base")
    )
    right = _scenario_doc(
        _scenario_series(sheet="Inputs", data_range="Inputs!B3:C3", scenario="Base")
    )
    merged = merge_series_binding_documents([left, right])
    series = merged["series"][0]
    assert series["sheet"] == "Inputs"
    assert series["data_range"] == ["Inputs!B2:C2", "Inputs!B3:C3"]
    assert series["structure"]["dimensions"][0]["bind"] == {"kind": "constant", "value": "Base"}


def test_merge_three_complementary_shards() -> None:
    docs = [
        _scenario_doc(_scenario_series(sheet=name, data_range=f"{name}!C12:X12", scenario=name))
        for name in ("Baseline", "B1", "B3")
    ]
    merged = merge_series_binding_documents(docs)
    series = merged["series"][0]
    assert series["sheet"] == ["Baseline", "B1", "B3"]
    assert series["data_range"] == ["Baseline!C12:X12", "B1!C12:X12", "B3!C12:X12"]
    assert series["structure"]["dimensions"][0]["bind"]["kind"] == "sheet_name"


def test_merge_rejects_different_keys() -> None:
    left = _scenario_doc(_scenario_series(sheet="Baseline", data_range="Baseline!C12:X12"))
    right = _scenario_doc(_scenario_series(sheet="B1", data_range="B1!C12:X12"))
    right["series"][0]["key"] = ["TIME_PERIOD"]
    with pytest.raises(SeriesBindingsLoadError, match="structural fields differ"):
        merge_series_binding_documents([left, right])


def test_merge_rejects_header_row_mismatch() -> None:
    left = _scenario_doc(_scenario_series(sheet="Baseline", data_range="Baseline!C12:X12"))
    right = _scenario_doc(_scenario_series(sheet="B1", data_range="B1!C12:X12"))
    right["series"][0]["structure"]["dimensions"][1]["bind"]["header_row"] = 1
    with pytest.raises(SeriesBindingsLoadError, match="structural fields differ"):
        merge_series_binding_documents([left, right])


def test_merge_rejects_input_and_output_across_different_sheets() -> None:
    left = _scenario_doc(
        _scenario_series(sheet="Baseline", data_range="Baseline!C12:X12", direction="input")
    )
    right = _scenario_doc(_scenario_series(sheet="B1", data_range="B1!C12:X12", direction="output"))
    with pytest.raises(SeriesBindingsLoadError, match="structural fields differ"):
        merge_series_binding_documents([left, right])


def test_merge_rejects_conflicting_constant_on_same_sheet() -> None:
    left = _scenario_doc(_scenario_series(sheet="Inputs", data_range="Inputs!B2:C2", scenario="A"))
    right = _scenario_doc(_scenario_series(sheet="Inputs", data_range="Inputs!B3:C3", scenario="B"))
    with pytest.raises(SeriesBindingsLoadError, match="structural fields differ|Cannot merge"):
        merge_series_binding_documents([left, right])


def _write_scenario_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    for sheet_name, values in (("Baseline", (1.0, 2.0)), ("B1", (10.0, 20.0))):
        ws = wb.add_worksheet(sheet_name)
        ws.write_number(7, 2, 2020)
        ws.write_number(7, 3, 2021)
        ws.write_number(11, 2, values[0])
        ws.write_number(11, 3, values[1])
    wb.close()


def test_resolve_concatenates_complementary_shard_cells(tmp_path: Path) -> None:
    wb_path = tmp_path / "scenarios.xlsx"
    _write_scenario_workbook(wb_path)
    left = _scenario_doc(
        _scenario_series(
            sheet="Baseline",
            data_range="Baseline!C12:D12",
            scenario="Baseline",
            direction="constant",
        )
    )
    right = _scenario_doc(
        _scenario_series(sheet="B1", data_range="B1!C12:D12", scenario="B1", direction="constant")
    )
    series = merge_series_binding_documents([left, right])["series"][0]
    series = validate_bindings_document({"schema_version": "1.14.0", "series": [series]})["series"][
        0
    ]

    targets = expand_data_range("Baseline!C12:D12") + expand_data_range("B1!C12:D12")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    resolved = resolve_series_binding(graph, wb_path, series, direction="constant")

    assert resolved["ok"] is True
    assert len(resolved["leaves"]) == 4
    by_key = {
        (leaf["key"]["SCENARIO"], leaf["key"]["TIME_PERIOD"]): leaf for leaf in resolved["leaves"]
    }
    assert set(by_key) == {("Baseline", 2020), ("Baseline", 2021), ("B1", 2020), ("B1", 2021)}
    assert by_key[("Baseline", 2020)]["address"] == "Baseline!C12"
    assert by_key[("B1", 2021)]["address"] == "B1!D12"


def test_resolve_sheet_name_values_map(tmp_path: Path) -> None:
    wb_path = tmp_path / "mapped.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    for sheet_name in ("Baseline DSA", "B1_GDP_ext"):
        ws = wb.add_worksheet(sheet_name)
        ws.write_number(7, 2, 2020)
        ws.write_number(11, 2, 1.0)
    wb.close()

    left = _scenario_doc(
        _scenario_series(
            sheet="Baseline DSA",
            data_range="Baseline DSA!C12",
            scenario="Baseline",
            direction="constant",
        )
    )
    right = _scenario_doc(
        _scenario_series(
            sheet="B1_GDP_ext",
            data_range="B1_GDP_ext!C12",
            scenario="B1",
            direction="constant",
        )
    )
    series = merge_series_binding_documents([left, right])["series"][0]
    series = validate_bindings_document({"schema_version": "1.14.0", "series": [series]})["series"][
        0
    ]

    targets = [
        *expand_data_range("'Baseline DSA'!C12"),
        *expand_data_range("B1_GDP_ext!C12"),
    ]
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    resolved = resolve_series_binding(graph, wb_path, series, direction="constant")
    assert resolved["ok"] is True
    scenarios = {leaf["key"]["SCENARIO"] for leaf in resolved["leaves"]}
    assert scenarios == {"Baseline", "B1"}


def test_resolve_authored_multi_range_sheet_name(tmp_path: Path) -> None:
    wb_path = tmp_path / "authored.xlsx"
    _write_scenario_workbook(wb_path)
    series = _scenario_series(
        sheet="Baseline",
        data_range="Baseline!C12:D12",
        bind={"kind": "sheet_name"},
        direction="constant",
    )
    series["sheet"] = ["Baseline", "B1"]
    series["data_range"] = ["Baseline!C12:D12", "B1!C12:D12"]
    bindings = validate_bindings_document(_scenario_doc(series))
    targets = expand_data_range("Baseline!C12:D12") + expand_data_range("B1!C12:D12")
    graph = create_dependency_graph(wb_path, targets, load_values=True)
    resolved = resolve_series_binding(graph, wb_path, bindings["series"][0], direction="constant")
    assert resolved["ok"] is True
    assert len(resolved["leaves"]) == 4
    assert {leaf["key"]["SCENARIO"] for leaf in resolved["leaves"]} == {"Baseline", "B1"}
