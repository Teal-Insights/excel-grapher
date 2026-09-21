"""Issue 868 — slotted Series.__post_init__ must not use zero-arg super()."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter import CodeGenerator
from excel_grapher.exporter.export_runtime.tensor import (
    Axis,
    Domain,
    SchemaError,
    Series,
    define_series,
)
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    inverted_graph_parts,
    invoke_public_compute,
    load_package,
    series_entry,
    write_workbook,
)


def test_define_series_with_values_constructs_slotted_series() -> None:
    axis = Axis("COMMODITY", ("Brent", "Copper"), str)
    series = define_series(
        "com_as_of_date",
        Domain.product(axis),
        (45345, 45345),
        cells={("Brent",): "COM!A4", ("Copper",): "COM!A5"},
        value_types=(int, bool, str, type(None)),
    )
    assert isinstance(series, Series)
    assert series[("Brent",)] == 45345
    assert series[("Copper",)] == 45345
    assert series.schema.series_id == "com_as_of_date"


def test_series_post_init_does_not_use_zero_arg_super() -> None:
    assert "__class__" not in Series.__post_init__.__code__.co_freevars


def test_series_post_init_still_runs_tensor_and_schema_checks() -> None:
    axis = Axis("COMMODITY", ("Brent", "Copper"), str)
    with pytest.raises(SchemaError, match="com_as_of_date"):
        define_series(
            "com_as_of_date",
            Domain.product(axis),
            (1.5, 2.5),
            cells={("Brent",): "COM!A4", ("Copper",): "COM!A5"},
            value_types=(int, bool, str, type(None)),
        )


def test_generate_modules_package_imports_valued_constant_series(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "const_out.xlsx",
        {
            "COM": {
                "A4": "Brent",
                "A5": "Copper",
                "B4": 45345,
                "B5": 45345,
                "C4": "=B4",
            }
        },
    )
    document = bindings_document(
        series_entry(
            "com_as_of_date",
            "COM!B4:B5",
            layout="series",
            direction="constant",
            dtype="int",
            label_column="A",
            key_concept="COMMODITY",
            key_read="string",
        ),
        series_entry("out", "COM!C4", layout="scalar", direction="output"),
    )
    _, _, graph = inverted_graph_parts(workbook, document)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document),
            bindings_workbook=workbook,
        )
    assert "define_series(" in modules["data.py"]
    package = load_package(modules, tmp_path, name="const_out")
    assert package.data.COM_AS_OF_DATE[("Brent",)] == 45345
    assert invoke_public_compute(package, package.compute_out, {}) == 45345
