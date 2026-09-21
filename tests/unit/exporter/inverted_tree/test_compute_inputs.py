"""Per-output frozen input dataclasses document each compute leaf closure."""

from __future__ import annotations

import dataclasses
import inspect
from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _two_cell_workbook(tmp_path: Path) -> Path:
    return write_workbook(tmp_path / "two_cell.xlsx", {"Inputs": {"A1": 2.0, "B1": "=A1*3"}})


def _two_cell_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("result", "Inputs!B1", layout="scalar", direction="output"),
    )


def _two_output_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "two_outputs.xlsx",
        {
            "Inputs": {"A1": 2.0, "A2": 5.0},
            "Outputs": {"A1": "=Inputs!A1*3", "A2": "=Inputs!A1+Inputs!A2"},
        },
    )


def _two_output_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("extra", "Inputs!A2", layout="scalar", direction="input"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
        series_entry("total", "Outputs!A2", layout="scalar", direction="output"),
    )


def test_result_inputs_from_defaults_computes_snapshot_and_override(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_two_cell_workbook(tmp_path), _two_cell_bindings()),
        tmp_path,
        name="compute_inputs",
    )
    inputs_cls = pkg.ResultInputs
    assert inputs_cls is pkg.api.ResultInputs
    assert inspect.signature(pkg.compute_result).parameters["inputs"].annotation in {
        inputs_cls,
        "ResultInputs",
    }
    assert list(inspect.signature(pkg.compute_result).parameters) == ["inputs"]
    snapshot = inputs_cls.from_defaults()
    assert pkg.compute_result(snapshot) == 6.0
    assert pkg.compute_result(inputs_cls.from_defaults(seed=4.0)) == 12.0
    with pytest.raises(TypeError):
        inputs_cls()
    with pytest.raises(TypeError, match="unknown"):
        inputs_cls.from_defaults(unknown=1)


def test_input_classes_are_exported_on_package_and_api(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_two_cell_workbook(tmp_path), _two_cell_bindings()),
        tmp_path,
        name="compute_inputs_export",
    )
    assert "ResultInputs" in pkg.api.__all__
    assert "ResultInputs" in pkg.__all__
    assert pkg.ResultInputs is pkg.api.ResultInputs


def test_sibling_input_classes_are_not_an_inheritance_tree(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_two_output_workbook(tmp_path), _two_output_bindings()),
        tmp_path,
        name="sibling_inputs",
    )
    assert pkg.ResultInputs is not pkg.TotalInputs
    assert not issubclass(pkg.ResultInputs, pkg.TotalInputs)
    assert not issubclass(pkg.TotalInputs, pkg.ResultInputs)
    result_fields = {field.name for field in dataclasses.fields(pkg.ResultInputs)}
    total_fields = {field.name for field in dataclasses.fields(pkg.TotalInputs)}
    assert result_fields == {"seed"}
    assert total_fields == {"extra", "seed"}
    assert pkg.compute_result(pkg.ResultInputs.from_defaults()) == 6.0
    assert pkg.compute_total(pkg.TotalInputs.from_defaults()) == 7.0
    assert pkg.compute_total(pkg.TotalInputs.from_defaults(seed=4.0, extra=1.0)) == 5.0
