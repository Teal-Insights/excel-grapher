"""Layer A5 — constants vs inputs in the signature."""

from __future__ import annotations

from pathlib import Path

import pytest

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


def test_public_constant_sets_live_in_data(tmp_path: Path) -> None:
    modules = generate_inverted(_a5_workbook(tmp_path), _a5_bindings())
    data, api = modules["data.py"], modules["api.py"]
    assert "_CONSTANTS_0 = frozenset({})\n" in data
    assert "_CONSTANTS_1 = frozenset({'engine_year_labels'})\n" in data
    assert "from .data import _CONSTANTS_0, _CONSTANTS_1\n" in api
    assert "_CONSTANTS_0 =" not in api
    assert "constants=_CONSTANTS_0," in api
    assert "constants=_CONSTANTS_1," in api


def test_year_labels_appear_only_on_shocked_compute(tmp_path: Path) -> None:
    workbook = _a5_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a5_bindings()), tmp_path, name="a5_pkg")
    baseline_all = all_param_names(pkg.compute_output_baseline)
    shocked_all = all_param_names(pkg.compute_output_shocked)
    assert "shock_year" not in baseline_all
    assert "engine_year_labels" not in baseline_all
    assert "engine_year_labels" not in shocked_all
    assert required_param_names(pkg.compute_output_baseline) == ("value",)
    assert "shock_year" in required_param_names(pkg.compute_output_shocked)
    assert "value" in required_param_names(pkg.compute_output_shocked)
    assert pkg.compute_output_baseline.__constants__ == ()
    assert pkg.compute_output_shocked.__constants__ == ("engine_year_labels",)
    assert "engine_year_labels" in all_param_names(pkg.internals.shocked_path)


def test_overriding_data_constant_changes_compute_and_restores(tmp_path: Path) -> None:
    workbook = _a5_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _a5_bindings()), tmp_path, name="a5_ov")
    baseline = pkg.compute_output_shocked(value=10.0, shock_year=1)
    assert (baseline[1], baseline[2]) == pytest.approx((11.0, 11.0))
    original = pkg.data.ENGINE_YEAR_LABELS
    zeros = pkg.data.EngineYearLabels.from_records(
        domain=original.domain, records=(((1,), 0), ((2,), 0))
    )
    pkg.data.ENGINE_YEAR_LABELS = zeros
    result = pkg.compute_output_shocked(value=10.0, shock_year=1)
    assert (result[1], result[2]) == pytest.approx((10.0, 10.0))
    pkg.data.ENGINE_YEAR_LABELS = original
    with pkg.data.overrides(ENGINE_YEAR_LABELS=zeros):
        result = pkg.compute_output_shocked(value=10.0, shock_year=1)
        assert (result[1], result[2]) == pytest.approx((10.0, 10.0))
    assert pkg.data.ENGINE_YEAR_LABELS is original
    result = pkg.compute_output_shocked(value=10.0, shock_year=1)
    assert (result[1], result[2]) == pytest.approx((11.0, 11.0))
    with (
        pytest.raises(AttributeError, match="unknown constant"),
        pkg.data.overrides(NOT_A_CONSTANT=(0, 0)),
    ):
        pass


def test_internal_constant_argument_validates_supplied_tensor(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_a5_workbook(tmp_path), _a5_bindings()), tmp_path, name="a5_schema"
    )
    wrong = pkg.Tensor.from_records(
        domain=pkg.Domain.product(pkg.Axis("TIME_PERIOD", (3, 4), int)),
        records=(((3,), 1), ((4,), 2)),
    )
    with pytest.raises(pkg.tensor.SchemaError, match="engine_year_labels.*required coordinate"):
        pkg.internals.shocked_path(value=10.0, shock_year=1, engine_year_labels=wrong)


def test_semantic_loop_does_not_overwrite_input_named_value(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_a5_workbook(tmp_path), _a5_bindings()), tmp_path, name="a5_value"
    )
    result = pkg.internals.shocked_path(
        value=10.0, shock_year=1, engine_year_labels=pkg.data.ENGINE_YEAR_LABELS
    )
    assert result[1] == 11.0
    assert result[2] == 11.0
