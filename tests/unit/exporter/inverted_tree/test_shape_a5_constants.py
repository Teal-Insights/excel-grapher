"""Layer A5 — constants vs inputs in the signature."""

from __future__ import annotations

from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)


def _a5_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a5.xlsx",
        {
            "Inputs": {
                "A1": 10.0,
                "B1": 2,
            },
            "Engine": {
                "C4": 1,
                "D4": 2,
                "C5": 1,
                "D5": 2,
                "C6": "=Inputs!A1",
                "D6": "=Inputs!A1",
                "C7": "=Inputs!A1+IF(Engine!C5>=Inputs!B1,1,0)",
                "D7": "=Inputs!A1+IF(Engine!D5>=Inputs!B1,1,0)",
            },
            "Outputs": {
                "A1": "=Engine!C6",
                "B1": "=Engine!D6",
                "A2": "=Engine!C7",
                "B2": "=Engine!D7",
                "A10": 1,
                "B10": 2,
            },
        },
    )


def _a5_bindings() -> dict:
    return bindings_document(
        series_entry("value", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("shock_year", "Inputs!B1", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "engine_year_labels",
            "Engine!C5:D5",
            layout="series",
            direction="constant",
            dtype="int",
            header_row=4,
        ),
        series_entry(
            "baseline_path",
            "Engine!C6:D6",
            layout="series",
            direction="internal",
            header_row=5,
        ),
        series_entry(
            "shocked_path",
            "Engine!C7:D7",
            layout="series",
            direction="internal",
            header_row=5,
        ),
        series_entry(
            "output_baseline",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
        series_entry(
            "output_shocked",
            "Outputs!A2:B2",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_year_labels_appear_only_on_shocked_compute(tmp_path: Path) -> None:
    workbook = _a5_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a5_bindings()), tmp_path, name="a5_pkg")
    baseline_all = all_param_names(pkg.compute_output_baseline)
    shocked_all = all_param_names(pkg.compute_output_shocked)
    assert "shock_year" not in baseline_all
    assert "engine_year_labels" not in baseline_all
    assert required_param_names(pkg.compute_output_baseline) == ("value",)
    assert "shock_year" in required_param_names(pkg.compute_output_shocked)
    assert "value" in required_param_names(pkg.compute_output_shocked)
    assert "engine_year_labels" in shocked_all
    assert "engine_year_labels" not in required_param_names(pkg.compute_output_shocked)
