"""Layer A7 — series members are measures: number or Excel error code."""

from __future__ import annotations

from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _mixed_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a7_mixed.xlsx",
        {
            "Inputs": {
                "B1": 10.0,
                "C1": 0.0,
                "B10": 1,
                "C10": 2,
            },
            "Engine": {
                "B1": "=1/Inputs!B1",
                "C1": "=1/Inputs!C1",
                "B10": 1,
                "C10": 2,
            },
            "Outputs": {
                "A1": "=Engine!B1",
                "B1": "=Engine!C1",
                "A10": 1,
                "B10": 2,
            },
        },
    )


def _mixed_bindings() -> dict:
    return bindings_document(
        series_entry(
            "denominators",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "engine_row",
            "Engine!B1:C1",
            layout="series",
            direction="internal",
            header_row=10,
        ),
        series_entry(
            "output_row",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_mixed_series_returns_error_code_not_abort(tmp_path: Path) -> None:
    workbook = _mixed_workbook(tmp_path)
    modules = generate_inverted(workbook, _mixed_bindings())
    assert "tuple[float, ...]" not in modules["api.py"]
    assert "float | str" in modules["api.py"]
    pkg = load_package(modules, tmp_path, name="a7_mixed")
    got = pkg.compute_output_row(denominators=(10.0, 0.0))
    assert got[0] == 0.1
    assert got[1] == "#DIV/0!"


def _ref_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a7_ref.xlsx",
        {
            "Engine": {"A1": "=#REF!"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )


def _ref_bindings() -> dict:
    return bindings_document(
        series_entry("engine_ref", "Engine!A1", layout="scalar", direction="internal"),
        series_entry("output_ref", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_ref_literal_is_error_code_measure(tmp_path: Path) -> None:
    workbook = _ref_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _ref_bindings()), tmp_path, name="a7_ref")
    assert pkg.compute_output_ref() == ("#REF!",)


def _scan_poison_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a7_scan.xlsx",
        {
            "Inputs": {
                "A1": 60.0,
                "B1": 0.0,
                "C1": 3.5,
                "B10": 1,
                "C10": 2,
            },
            "Engine": {
                "B1": "=Inputs!A1",
                "C1": "=B1/Inputs!B1",
                "D1": "=C1/Inputs!C1",
                "C10": 1,
                "D10": 2,
            },
            "Outputs": {
                "A1": "=Engine!C1",
                "B1": "=Engine!D1",
                "A10": 1,
                "B10": 2,
            },
        },
    )


def _scan_poison_bindings() -> dict:
    return bindings_document(
        series_entry("initial_debt", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "growth",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry("engine_year0", "Engine!B1", layout="scalar", direction="internal"),
        series_entry(
            "engine_path",
            "Engine!C1:D1",
            layout="series",
            direction="internal",
            header_row=10,
        ),
        series_entry(
            "output_path",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_scan_propagates_error_code_to_later_years(tmp_path: Path) -> None:
    workbook = _scan_poison_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _scan_poison_bindings()),
        tmp_path,
        name="a7_scan",
    )
    got = pkg.compute_output_path(initial_debt=60.0, growth=(0.0, 3.5))
    assert got == ("#DIV/0!", "#DIV/0!")
