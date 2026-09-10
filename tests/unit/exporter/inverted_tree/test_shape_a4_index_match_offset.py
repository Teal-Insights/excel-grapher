"""Layer A4 — INDEX/MATCH and OFFSET become indexing."""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import pytest

from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _index_match_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a4_index.xlsx",
        {
            "Inputs": {
                "B5": "Litellia",
                "A10": "Borvelia",
                "B10": 60.0,
                "C10": "x",
                "A11": "Litellia",
                "B11": 80.0,
                "C11": "x",
                "A12": "Aurelium",
                "B12": 40.0,
                "C12": "x",
            },
            "Engine": {
                "B6": "=INDEX(Inputs!$A$10:$C$12,MATCH(Inputs!$B$5,Inputs!$A$10:$A$12,0),2)",
            },
            "Outputs": {
                "A1": "=Engine!B6",
            },
        },
    )


def _index_match_bindings() -> dict:
    return bindings_document(
        series_entry(
            "country_name",
            "Inputs!B5",
            layout="scalar",
            direction="input",
            dtype="string",
        ),
        series_entry(
            "country_profile_names",
            "Inputs!A10:A12",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "country_initial_debt",
            "Inputs!B10:B12",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "initial_debt_resolved",
            "Engine!B6",
            layout="scalar",
            direction="internal",
        ),
        series_entry("output_debt", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_index_match_has_no_xl_index_ref(tmp_path: Path) -> None:
    workbook = _index_match_workbook(tmp_path)
    modules = generate_inverted(workbook, _index_match_bindings())
    internals = modules["internals.py"]
    assert "xl_index_ref" not in internals
    assert "xl_cell" not in internals
    assert "ctx" not in internals
    pkg = load_package(modules, tmp_path, name="a4_idx")
    params = all_param_names(pkg.internals.initial_debt_resolved)
    assert "country_name" in params
    assert "country_profile_names" in params
    assert "country_initial_debt" in params
    assert pkg.internals.initial_debt_resolved(
        country_name="Litellia",
        country_profile_names=pkg.data.COUNTRY_PROFILE_NAMES,
        country_initial_debt=pkg.data.COUNTRY_INITIAL_DEBT_DEFAULT,
    ) == pytest.approx(80.0)


def test_unknown_match_returns_na(tmp_path: Path) -> None:
    workbook = _index_match_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _index_match_bindings()), tmp_path, name="a4_na")
    assert (
        pkg.internals.initial_debt_resolved(
            country_name="NotACountry",
            country_profile_names=pkg.data.COUNTRY_PROFILE_NAMES,
            country_initial_debt=pkg.data.COUNTRY_INITIAL_DEBT_DEFAULT,
        )
        == "#N/A"
    )


def _offset_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a4_offset.xlsx",
        {
            "Inputs": {
                "A1": 2,
                "B1": -2.0,
                "C1": 2.0,
                "D1": -1.0,
                "B10": 1,
                "C10": 2,
                "D10": 3,
            },
            "Engine": {
                "A1": "=OFFSET(Inputs!B1,0,Inputs!A1-1)",
            },
            "Outputs": {
                "A1": "=Engine!A1",
            },
        },
    )


def _offset_bindings() -> dict:
    return bindings_document(
        series_entry("shock_type", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "shock_magnitudes",
            "Inputs!B1:D1",
            layout="series",
            direction="input",
            header_row=10,
            key_concept="SHOCK_PARAMETER",
            key_read="int",
        ),
        series_entry(
            "shock_magnitude_resolved",
            "Engine!A1",
            layout="scalar",
            direction="internal",
        ),
        series_entry("output_mag", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_offset_into_row_is_indexing(tmp_path: Path) -> None:
    workbook = _offset_workbook(tmp_path)
    modules = generate_inverted(
        workbook,
        _offset_bindings(),
        dynamic_refs=DynamicRefConfig.from_constraints({"Inputs!A1": Literal[1, 2, 3]}, {}),
    )
    assert "xl_offset" not in modules["internals.py"]
    assert "shock_magnitudes: data.ShockMagnitudes" in modules["internals.py"]
    assert "(_ for _ in ())" not in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="a4_off")
    resolved = pkg.internals.shock_magnitude_resolved
    values = pkg.data.SHOCK_MAGNITUDES_DEFAULT
    assert resolved(shock_type=1, shock_magnitudes=values) == pytest.approx(-2.0)
    assert resolved(shock_type=2, shock_magnitudes=values) == pytest.approx(2.0)
    assert resolved(shock_type=3, shock_magnitudes=values) == pytest.approx(-1.0)
    assert resolved(shock_type=4, shock_magnitudes=values) == "#VALUE!"


def test_index_keeps_off_catalog_blanks_in_selected_column(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "index_blank.xlsx",
        {
            "Model": {
                "B1": 10,
                "B3": 30,
                "D1": "=INDEX(A1:B3,1,2)",
            }
        },
    )
    document = bindings_document(
        series_entry("first", "Model!B1", direction="input"),
        series_entry("last", "Model!B3", direction="constant"),
        series_entry("out", "Model!D1", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    pkg = load_package(modules, tmp_path, name="a4_blank_column")
    assert pkg.compute_out(first=10.0) == 10.0
