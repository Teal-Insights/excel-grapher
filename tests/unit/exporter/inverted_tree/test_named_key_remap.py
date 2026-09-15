"""Same-axis string remaps fold to a dict instead of a host-key if-ladder (#852).

When host and producer share an axis name but not string identity (scenario
codes vs long labels, ISO3 vs country names), each cell is a direct reference
whose body differs only by that producer string. Family merge already
parameterizes the other drivers; the leftover ladder is the 1:1 map.

A producer string that is not a function of the host key (the same host
key mapping onto two producer labels) stays a literal. Extra predicates on
some members stay a leftover family beside the remapped path. A many-to-one
map is still a function of the host key and still folds.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _dim_row(field: str, column: str) -> dict[str, Any]:
    return {
        "id": field,
        "concept": field,
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": column, "read": "string"},
    }


def _dim_col(header_row: int) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _grid_series(
    series_id: str,
    data_range: str,
    key: list[str],
    dims: list[dict[str, Any]],
    *,
    kind: str = "input",
    compute: str | None = None,
) -> dict[str, Any]:
    block: dict[str, Any] = {
        "id": series_id,
        "sheet": data_range.split("!", 1)[0],
        "data_range": data_range,
        "layout": "series",
        "key": key,
        "structure": {"measure": _measure(), "dimensions": dims},
    }
    if kind == "input":
        block["input"] = {}
    else:
        block["output"] = {"compute": {"name": compute or f"compute_{series_id}"}}
    return block


def _with_concepts(document: dict[str, Any], *fields: str) -> dict[str, Any]:
    known = {item["id"] for item in document["concept_scheme"]["concepts"]}
    for field in fields:
        if field not in known:
            document["concept_scheme"]["concepts"].append({"id": field, "dtype": "string"})
    return document


def _scenario_remap_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "scenario_label.xlsx",
        {
            "Data": {
                "A1": "scenario",
                "B1": 2025,
                "C1": 2026,
                "A2": "Bounds Test 1: Real GDP Growth Shock",
                "B2": 2,
                "C2": 12,
                "A3": "Bound Test 2: Exports Shock",
                "B3": 3,
                "C3": 13,
                "A6": "scenario",
                "B6": 2025,
                "C6": 2026,
                "A7": "B1",
                "B7": "=B2",
                "C7": "=C2",
                "A8": "B3",
                "B8": "=B3",
                "C8": "=C3",
            }
        },
    )


def _scenario_remap_bindings() -> dict[str, Any]:
    return _with_concepts(
        bindings_document(
            _grid_series(
                "stress_interest",
                "Data!B2:C3",
                ["SCENARIO", "TIME_PERIOD"],
                [_dim_row("SCENARIO", "A"), _dim_col(1)],
            ),
            _grid_series(
                "dsa_interest",
                "Data!B7:C8",
                ["SCENARIO", "TIME_PERIOD"],
                [_dim_row("SCENARIO", "A"), _dim_col(6)],
                kind="output",
                compute="compute_dsa_interest",
            ),
            schema_version="1.16.0",
        ),
        "SCENARIO",
    )


def _iso3_remap_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "iso3_remap.xlsx",
        {
            "Data": {
                "A1": "country",
                "B1": "ifs",
                "A2": "Bangladesh",
                "B2": 513,
                "A3": "Benin",
                "B3": 638,
                "A4": "Bhutan",
                "B4": 514,
                "D1": "iso3",
                "E1": "ifs",
                "D2": "BGD",
                "E2": "=B2",
                "D3": "BEN",
                "E3": "=B3",
                "D4": "BTN",
                "E4": "=B4",
            }
        },
    )


def _iso3_remap_bindings() -> dict[str, Any]:
    return _with_concepts(
        bindings_document(
            _grid_series(
                "catalog_ifs",
                "Data!B2:B4",
                ["REF_AREA"],
                [_dim_row("REF_AREA", "A")],
            ),
            _grid_series(
                "trigger_ifs",
                "Data!E2:E4",
                ["REF_AREA"],
                [_dim_row("REF_AREA", "D")],
                kind="output",
                compute="compute_trigger_ifs",
            ),
            schema_version="1.16.0",
        ),
        "REF_AREA",
    )


def test_scenario_code_to_label_remap_is_a_dict_not_an_if_ladder(tmp_path: Path) -> None:
    workbook = _scenario_remap_workbook(tmp_path)
    document = _scenario_remap_bindings()
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "SCENARIO_KEY" in internals
    assert "SCENARIO_KEY[scenario]" in internals
    assert "stress_interest[SCENARIO_KEY[scenario], time_period]" in internals
    assert not re.search(r"if scenario ==", internals)
    assert "Bounds Test 1: Real GDP Growth Shock" in internals
    assert "'B1'" in internals
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "scenario_remap")
    got = pkg.compute_dsa_interest(stress_interest=pkg.data.STRESS_INTEREST_DEFAULT)
    assert got["B1", 2025] == 2.0
    assert got["B3", 2026] == 13.0


def test_iso3_to_name_remap_is_a_dict_not_an_if_ladder(tmp_path: Path) -> None:
    workbook = _iso3_remap_workbook(tmp_path)
    document = _iso3_remap_bindings()
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "REF_AREA_KEY" in internals
    assert "REF_AREA_KEY[ref_area]" in internals
    assert "catalog_ifs[REF_AREA_KEY[ref_area]]" in internals
    assert not re.search(r"if ref_area ==", internals)
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "iso3_remap")
    got = pkg.compute_trigger_ifs(catalog_ifs=pkg.data.CATALOG_IFS_DEFAULT)
    assert got["BGD"] == 513.0
    assert got["BTN"] == 514.0


def test_inverse_label_to_code_remap_uses_the_same_fold(tmp_path: Path) -> None:
    """Chart-side long labels onto DSA codes are the same 1:1 map, inverted."""
    workbook = write_workbook(
        tmp_path / "scenario_inverse.xlsx",
        {
            "Data": {
                "A1": "scenario",
                "B1": 2025,
                "C1": 2026,
                "A2": "B1",
                "B2": 2,
                "C2": 12,
                "A3": "B3",
                "B3": 3,
                "C3": 13,
                "A6": "scenario",
                "B6": 2025,
                "C6": 2026,
                "A7": "Bounds Test 1: Real GDP Growth Shock",
                "B7": "=B2",
                "C7": "=C2",
                "A8": "Bound Test 2: Exports Shock",
                "B8": "=B3",
                "C8": "=C3",
            }
        },
    )
    document = _with_concepts(
        bindings_document(
            _grid_series(
                "dsa_interest",
                "Data!B2:C3",
                ["SCENARIO", "TIME_PERIOD"],
                [_dim_row("SCENARIO", "A"), _dim_col(1)],
            ),
            _grid_series(
                "chart_interest",
                "Data!B7:C8",
                ["SCENARIO", "TIME_PERIOD"],
                [_dim_row("SCENARIO", "A"), _dim_col(6)],
                kind="output",
                compute="compute_chart_interest",
            ),
            schema_version="1.16.0",
        ),
        "SCENARIO",
    )
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "SCENARIO_KEY[scenario]" in internals
    assert not re.search(r"if scenario ==", internals)
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "scenario_inverse")
    got = pkg.compute_chart_interest(dsa_interest=pkg.data.DSA_INTEREST_DEFAULT)
    assert got["Bounds Test 1: Real GDP Growth Shock", 2025] == 2.0
    assert got["Bound Test 2: Exports Shock", 2026] == 13.0


def test_extra_predicate_member_stays_a_leftover_family(tmp_path: Path) -> None:
    """A member whose body is not a pure remap keeps its own family (#852)."""
    workbook = write_workbook(
        tmp_path / "scenario_leftover.xlsx",
        {
            "Data": {
                "A1": "scenario",
                "B1": 2025,
                "C1": 2026,
                "A2": "Bounds Test 1: Real GDP Growth Shock",
                "B2": 2,
                "C2": 12,
                "A3": "Bound Test 2: Exports Shock",
                "B3": 3,
                "C3": 13,
                "A4": "Market Borrowing Shock",
                "B4": 4,
                "C4": 14,
                "A7": "scenario",
                "B7": 2025,
                "C7": 2026,
                "A8": "B1",
                "B8": "=B2",
                "C8": "=C2",
                "A9": "B3",
                "B9": "=B3",
                "C9": "=C3",
                "A10": "B6",
                "B10": "=B4*2",
                "C10": "=C4*2",
            }
        },
    )
    document = _with_concepts(
        bindings_document(
            _grid_series(
                "stress_interest",
                "Data!B2:C4",
                ["SCENARIO", "TIME_PERIOD"],
                [_dim_row("SCENARIO", "A"), _dim_col(1)],
            ),
            _grid_series(
                "dsa_interest",
                "Data!B8:C10",
                ["SCENARIO", "TIME_PERIOD"],
                [_dim_row("SCENARIO", "A"), _dim_col(7)],
                kind="output",
                compute="compute_dsa_interest",
            ),
            schema_version="1.16.0",
        ),
        "SCENARIO",
    )
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "SCENARIO_KEY[scenario]" in internals
    assert re.search(r"if scenario == 'B6':", internals)
    assert "xl_mul(" in internals
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "scenario_leftover")
    got = pkg.compute_dsa_interest(stress_interest=pkg.data.STRESS_INTEREST_DEFAULT)
    assert got["B1", 2025] == 2.0
    assert got["B6", 2026] == 28.0


def test_many_to_one_string_map_still_folds(tmp_path: Path) -> None:
    """A function of the host key folds even when it is not injective."""
    workbook = write_workbook(
        tmp_path / "iso3_many_to_one.xlsx",
        {
            "Data": {
                "A1": "country",
                "B1": "ifs",
                "A2": "Bangladesh",
                "B2": 513,
                "A3": "Benin",
                "B3": 638,
                "A4": "Bhutan",
                "B4": 514,
                "D1": "iso3",
                "E1": "ifs",
                "D2": "BGD",
                "E2": "=B3",
                "D3": "BEN",
                "E3": "=B2",
                "D4": "BTN",
                "E4": "=B2",
            }
        },
    )
    document = _iso3_remap_bindings()
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "REF_AREA_KEY[ref_area]" in internals
    assert not re.search(r"if ref_area ==", internals)
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "iso3_many_to_one")
    got = pkg.compute_trigger_ifs(catalog_ifs=pkg.data.CATALOG_IFS_DEFAULT)
    assert got["BGD"] == 638.0
    assert got["BEN"] == 513.0
    assert got["BTN"] == 513.0


def test_producer_key_that_is_not_a_function_of_the_host_stays_literal(tmp_path: Path) -> None:
    """The same host key mapping to two producer strings is not a remap."""
    workbook = write_workbook(
        tmp_path / "iso3_independent.xlsx",
        {
            "Data": {
                "A1": "country",
                "B1": 2025,
                "C1": 2026,
                "A2": "Bangladesh",
                "B2": 1,
                "C2": 2,
                "A3": "Benin",
                "B3": 3,
                "C3": 4,
                "D1": "iso3",
                "E1": 2025,
                "F1": 2026,
                "D2": "BGD",
                "E2": "=B2",
                "F2": "=C3",
                "D3": "BEN",
                "E3": "=B3",
                "F3": "=C2",
            }
        },
    )
    document = _with_concepts(
        bindings_document(
            _grid_series(
                "catalog_ifs",
                "Data!B2:C3",
                ["REF_AREA", "TIME_PERIOD"],
                [_dim_row("REF_AREA", "A"), _dim_col(1)],
            ),
            _grid_series(
                "trigger_ifs",
                "Data!E2:F3",
                ["REF_AREA", "TIME_PERIOD"],
                [_dim_row("REF_AREA", "D"), _dim_col(1)],
                kind="output",
                compute="compute_trigger_ifs",
            ),
            schema_version="1.16.0",
        ),
        "REF_AREA",
    )
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "REF_AREA_KEY" not in internals
    assert "catalog_ifs['Bangladesh', time_period]" in internals
    assert "catalog_ifs['Benin', time_period]" in internals
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "iso3_independent")
    got = pkg.compute_trigger_ifs(catalog_ifs=pkg.data.CATALOG_IFS_DEFAULT)
    assert got["BGD", 2025] == 1.0
    assert got["BGD", 2026] == 4.0
    assert got["BEN", 2025] == 3.0
    assert got["BEN", 2026] == 2.0


def test_mismatched_axis_names_do_not_invent_a_remap(tmp_path: Path) -> None:
    """Cross-axis neighbor reads keep literals; only a shared axis name remaps."""
    workbook = write_workbook(
        tmp_path / "cross_axis.xlsx",
        {
            "Data": {
                "A2": "1 Year",
                "A3": "2 Year",
                "A4": "10 Year",
                "B2": "=A3",
                "B3": "=A4",
                "B4": "=A2",
            }
        },
    )
    document = bindings_document(
        series_entry(
            "labels",
            "Data!A2:A4",
            layout="series",
            direction="input",
            dtype="string",
            label_column="A",
            key_concept="VARIANT",
            key_read="string",
        ),
        series_entry(
            "picked",
            "Data!B2:B4",
            layout="series",
            direction="output",
            dtype="string",
            label_column="A",
            key_concept="TENOR",
            key_read="string",
        ),
        schema_version="1.16.0",
    )
    for field in ("VARIANT", "TENOR"):
        if field not in {concept["id"] for concept in document["concept_scheme"]["concepts"]}:
            document["concept_scheme"]["concepts"].append({"id": field, "dtype": "string"})
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "TENOR_KEY" not in internals
    assert "VARIANT_KEY" not in internals
    assert "labels['2 Year']" in internals
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "cross_axis_literal")
    got = pkg.compute_picked(labels=pkg.data.LABELS_DEFAULT)
    assert got["1 Year"] == "2 Year"
    assert got["10 Year"] == "1 Year"
