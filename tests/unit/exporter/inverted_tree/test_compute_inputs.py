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
from tests.unit.exporter.inverted_tree.test_input_domain import (
    _enum_flag_bindings,
    _enum_flag_workbook,
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


def _constant_output_workbook(tmp_path: Path) -> Path:
    return write_workbook(tmp_path / "constant_output.xlsx", {"Sheet": {"A1": 5.0, "B1": "=A1*2"}})


def _constant_output_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Sheet!A1", layout="scalar", direction="constant"),
        series_entry("result", "Sheet!B1", layout="scalar", direction="output"),
    )


def test_empty_inputs_still_require_the_bundle(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_constant_output_workbook(tmp_path), _constant_output_bindings()),
        tmp_path,
        name="empty_inputs",
    )
    inputs_cls = pkg.ResultInputs
    assert dataclasses.fields(inputs_cls) == ()
    with pytest.raises(TypeError, match="missing 1 required positional argument"):
        pkg.compute_result()
    assert pkg.compute_result(inputs_cls.from_defaults()) == 10.0
    assert pkg.compute_result(inputs_cls()) == 10.0


def test_result_inputs_from_defaults_computes_snapshot_and_override(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_two_cell_workbook(tmp_path), _two_cell_bindings()),
        tmp_path,
        name="compute_inputs",
    )
    inputs_cls = pkg.ResultInputs
    assert inputs_cls is pkg.api.ResultInputs
    assert inputs_cls is pkg.model.ResultInputs
    assert inspect.signature(pkg.compute_result).parameters["inputs"].annotation in {
        inputs_cls,
        "ResultInputs",
    }
    assert list(inspect.signature(pkg.compute_result).parameters) == ["inputs"]
    assert inputs_cls.from_defaults.__annotations__["return"] in {inputs_cls, "Self"}
    snapshot = inputs_cls.from_defaults()
    assert type(snapshot) is inputs_cls
    assert pkg.compute_result(snapshot) == 6.0
    assert pkg.compute_result(inputs_cls.from_defaults(seed=4.0)) == 12.0
    with pytest.raises(TypeError):
        inputs_cls()
    with pytest.raises(TypeError):
        inputs_cls(4.0)
    with pytest.raises(TypeError, match="unknown"):
        inputs_cls.from_defaults(unknown=1)


def test_input_classes_are_exported_on_package_and_api(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_two_cell_workbook(tmp_path), _two_cell_bindings()),
        tmp_path,
        name="compute_inputs_export",
    )
    assert "ResultInputs" in pkg.api.__all__
    assert "ResultInputs" in pkg.model.__all__
    assert "ResultInputs" in pkg.__all__
    assert pkg.ResultInputs is pkg.api.ResultInputs
    assert pkg.ResultInputs is pkg.model.ResultInputs
    assert "Model" not in pkg.api.__all__
    assert not hasattr(pkg.api, "Model")


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


def test_compute_rejects_the_wrong_inputs_class(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_two_output_workbook(tmp_path), _two_output_bindings()),
        tmp_path,
        name="wrong_inputs",
    )
    with pytest.raises(TypeError, match="compute_total\\(\\) expected TotalInputs"):
        pkg.compute_total(pkg.ResultInputs.from_defaults())


def test_checks_run_at_inputs_construction(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="inputs_domain",
    )
    assert pkg.compute_out(pkg.OutInputs(flag=0)) == 0
    assert pkg.compute_out(pkg.OutInputs.from_defaults()) == 0
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.OutInputs(flag=2)
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.OutInputs.from_defaults(flag=2)


def test_model_skips_checks_for_a_validated_bundle(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="model_skip",
    )
    valid = pkg.OutInputs(flag=1)
    object.__setattr__(valid, "flag", 2)
    assert pkg.Model(valid).out == 2
    with pytest.raises(TypeError, match="bound inputs instance"):
        pkg.Model({"flag": 1})
    with pytest.raises(TypeError, match="keyword inputs"):
        pkg.Model(valid, flag=0)
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.Model(flag=2)
