"""Issue 1006 — generated `*Inputs` records expose each field's domain."""

from __future__ import annotations

import dataclasses
from pathlib import Path
from typing import Annotated, Literal, get_args

import pytest

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.series_bindings.load import load_series_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.paths import INVERTED_TREE_TINY_DSA, INVERTED_TREE_TINY_DSA_LABELLED
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _tiny_dsa(root: Path, tmp_path: Path, name: str):
    workbook = next(root.glob("*.xlsx"))
    bindings_dir = root / "bindings"
    bindings = load_series_bindings(bindings_dir)
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_bindings(bindings, workbook, bindings_path=bindings_dir),
    )
    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(series_bindings=bindings, bindings_workbook=workbook)
    return load_package(modules, tmp_path, name=name)


@pytest.fixture(scope="module")
def tiny_dsa(tmp_path_factory: pytest.TempPathFactory):
    return _tiny_dsa(INVERTED_TREE_TINY_DSA, tmp_path_factory.mktemp("describe"), "describe_dsa")


def test_describe_lists_every_field_in_declaration_order(tiny_dsa) -> None:
    described = tiny_dsa.OutputShockedInputs.describe()
    names = [field.name for field in dataclasses.fields(tiny_dsa.OutputShockedInputs)]
    assert list(described) == names
    assert all(isinstance(item, tiny_dsa.InputField) for item in described.values())
    assert all(described[name].name == name for name in names)


def test_describe_series_field_reports_axes_keys_bounds_and_provenance(tiny_dsa) -> None:
    data = tiny_dsa.data
    field = tiny_dsa.OutputShockedInputs.describe()["interest_baseline"]
    assert field.is_series
    assert field.size == 5
    assert [(axis.name, axis.keys) for axis in field.axes] == [("TIME_PERIOD", (1, 2, 3, 4, 5))]
    assert field.domain == data.INTEREST_BASELINE.domain
    assert field.default is data.INTEREST_BASELINE_DEFAULT
    assert field.cells is data.INTEREST_BASELINE.cells
    element, none = get_args(field.values)
    assert none is type(None)
    assert get_args(element)[0] is float
    assert get_args(element)[1] == tiny_dsa.runtime.RealBetween(0.0, 20.0)

    shocks = tiny_dsa.OutputShockedInputs.describe()["shock_magnitudes"]
    assert [(axis.name, axis.keys) for axis in shocks.axes] == [
        ("SHOCK_PARAMETER", ("Growth", "Interest", "Primary balance"))
    ]


def test_describe_scalar_field_has_no_axes(tiny_dsa) -> None:
    described = tiny_dsa.OutputShockedInputs.describe()
    year = described["shock_year"]
    assert not year.is_series
    assert year.domain is None
    assert year.axes == ()
    assert year.size == 1
    assert year.default == 2
    assert year.cells == {(): "Inputs!B21"}
    assert year.values == Annotated[int, tiny_dsa.runtime.Between(1, 5)]
    assert described["shock_type"].values == Literal[1, 2, 3]


def test_fields_link_to_data_specs_in_metadata(tiny_dsa) -> None:
    data = tiny_dsa.data
    by_name = {f.name: f for f in dataclasses.fields(tiny_dsa.OutputShockedInputs)}
    assert by_name["interest_baseline"].metadata["default"] is data.INTEREST_BASELINE_DEFAULT
    assert by_name["interest_baseline"].metadata["cells"] is data.INTEREST_BASELINE.cells
    assert by_name["shock_year"].metadata["default"] == data.SHOCK_YEAR_DEFAULT
    assert by_name["shock_year"].metadata["cells"] is data.SHOCK_YEAR_CELLS


def test_input_field_is_public(tiny_dsa) -> None:
    assert "InputField" in tiny_dsa.__all__
    assert tiny_dsa.InputField is tiny_dsa.runtime.InputField


def test_describe_labelled_axis_reports_snapshot_keys(tmp_path: Path) -> None:
    pkg = _tiny_dsa(INVERTED_TREE_TINY_DSA_LABELLED, tmp_path, "describe_labelled")
    field = pkg.OutputBaselineInputs.describe()["growth_baseline"]
    assert [(axis.name, axis.keys) for axis in field.axes] == [("TIME_PERIOD", (1, 2, 3, 4, 5))]
    assert field.default is pkg.data.GROWTH_BASELINE_DEFAULT


def test_inputs_class_source_snapshot(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "snap.xlsx",
        {
            "Inputs": {"A1": 2.0, "B1": 0.25, "C1": 0.5, "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!A1*Inputs!B1", "B1": "=Inputs!C1", "A10": 1, "B10": 2},
        },
    )
    document = bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "rate",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
            domain={"real_between": {"min": 0, "max": 1}},
        ),
        series_entry("out", "Outputs!A1:B1", layout="series", direction="output", header_row=10),
    )
    source = generate_inverted(workbook, document)["model.py"]
    start = source.index("@dataclass(frozen=True, kw_only=True)\nclass OutInputs")
    end = source.index("\n\n\n", start)
    assert source[start:end] == (
        "@dataclass(frozen=True, kw_only=True)\n"
        "class OutInputs(_SnapshotInputs):\n"
        '    """Bound input leaves for `compute_out`."""\n'
        "\n"
        "    seed: float | str = _field(\n"
        '        metadata={"default": data.SEED_DEFAULT, "cells": data.SEED_CELLS},\n'
        "    )\n"
        "    rate: data.Rate = _field(\n"
        '        metadata={"default": data.RATE_DEFAULT, "cells": data.RATE.cells},\n'
        "    )"
    )


def test_describe_rejects_an_unparameterized_series_annotation(tiny_dsa) -> None:
    @dataclasses.dataclass(frozen=True)
    class Bare:
        growth: object = dataclasses.field(
            metadata={"default": tiny_dsa.data.GROWTH_BASELINE_DEFAULT, "cells": None}
        )

    with pytest.raises(TypeError, match=r"growth: series annotation .* is not parameterized"):
        tiny_dsa.runtime.describe_inputs(Bare)
