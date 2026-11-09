"""Domain constructors and `span` are imported from the emitter, not source text."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Literal

import pytest

from excel_grapher.core.cell_types import Between, RealBetween
from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)

_MODULES = ("data.py", "internals.py", "validation.py", "model.py", "api.py")
_DOMAIN_NAMES = frozenset({"Annotated", "Between", "Literal", "RealBetween"})
_TEXT = {
    "between": "Between(0, 1)",
    "literal": "Literal['a']",
    "annotated": "Annotated[int, Between(0, 1)]",
    "real_between": "RealBetween(0, 1)",
    "span": "span(0, 1)",
}
_IMPORT_NAME = {
    "between": "Between",
    "literal": "Literal",
    "annotated": "Annotated",
    "real_between": "RealBetween",
    "span": "span",
}


def _imported(source: str) -> set[str]:
    """Names bound by `from typing` and `from .runtime` import lines."""
    names: set[str] = set()
    lines = source.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index].strip()
        if line.startswith("from typing import ") or line.startswith("from .runtime import "):
            blob = line
            while blob.endswith("\\") or ("(" in blob and ")" not in blob):
                index += 1
                blob = blob.rstrip("\\") + " " + lines[index].strip()
            payload = blob.split(" import ", 1)[1].replace("(", " ").replace(")", " ")
            names.update(part.strip() for part in payload.split(",") if part.strip())
        index += 1
    return names


def _domain_imports(source: str) -> set[str]:
    return _imported(source) & _DOMAIN_NAMES


def _series(series_id: str, data_range: str, *, direction: str, dtype: str) -> dict[str, object]:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction=direction,
        dtype=dtype,
        label_column="A",
        key_concept="COUNTRY",
        key_read="string",
    )


def _export_text(tmp_path: Path, series_id: str, text: str, *, where: str) -> dict[str, str]:
    if where == "cell":
        cells: dict[str, object] = {
            "A2": "grace",
            "A3": "maturity",
            "C2": text,
            "C3": "plain",
            "I2": "=C2",
            "I3": "=C3",
        }
        dtype = "string"
    else:
        cells = {
            "A2": text,
            "A3": "other",
            "C2": 1,
            "C3": 2,
            "I2": "=C2",
            "I3": "=C3",
        }
        dtype = "float"
    workbook = write_workbook(tmp_path / "book.xlsx", {"Inputs": cells})
    document = bindings_document(
        _series(series_id, "Inputs!C2:C3", direction="input", dtype=dtype),
        _series("rate_out", "Inputs!I2:I3", direction="output", dtype=dtype),
        schema_version="1.19.0",
    )
    return generate_inverted(workbook, document)


@pytest.mark.parametrize("series_id", list(_TEXT))
@pytest.mark.parametrize("where", ["cell", "axis"])
def test_constructor_text_in_a_cell_or_axis_key_is_not_an_import(
    tmp_path: Path, series_id: str, where: str
) -> None:
    """A workbook string that looks like a constructor does not reserve that name."""
    text = _TEXT[series_id]
    modules = _export_text(tmp_path, series_id, text, where=where)
    assert text in modules["data.py"]
    imported = _IMPORT_NAME[series_id]
    for name in _MODULES:
        assert imported not in _imported(modules[name])
    load_package(modules, tmp_path, name=f"text_{series_id}_{where}")


@pytest.mark.parametrize(
    ("series_id", "dtype", "values", "constraints", "imported"),
    [
        (
            "between",
            "int",
            (0, 1),
            {
                "Inputs!C2": Annotated[int, Between(0, 1)],
                "Inputs!C3": Annotated[int, Between(0, 1)],
            },
            "Between",
        ),
        (
            "literal",
            "string",
            ("On", "Off"),
            {
                "Inputs!C2": Literal["Off", "On"],
                "Inputs!C3": Literal["Off", "On"],
            },
            "Literal",
        ),
    ],
)
def test_real_annotation_still_collides_with_the_public_alias(
    tmp_path: Path,
    series_id: str,
    dtype: str,
    values: tuple[object, object],
    constraints: dict[str, object],
    imported: str,
) -> None:
    """A spliced constructor still fails closed when the alias uses that name."""
    left, right = values
    workbook = write_workbook(
        tmp_path / "book.xlsx",
        {
            "Inputs": {
                "A2": "grace",
                "A3": "maturity",
                "C2": left,
                "C3": right,
                "I2": "=C2",
                "I3": "=C3",
            }
        },
    )
    document = bindings_document(
        _series(series_id, "Inputs!C2:C3", direction="input", dtype=dtype),
        _series("rate_out", "Inputs!I2:I3", direction="output", dtype=dtype),
        schema_version="1.19.0",
    )
    with pytest.raises(InvertedTreeExportError, match=f"generated name {imported}") as raised:
        generate_inverted(
            workbook,
            document,
            dynamic_refs=DynamicRefConfig.from_constraints(constraints),
        )
    message = str(raised.value)
    assert f"import {imported}" in message
    assert f"series {series_id!r} alias" in message


def _scalar_pair(
    tmp_path: Path,
    *,
    series_id: str,
    dtype: str,
    value: object,
    output_dtype: str,
    formula: str,
    constraint: object,
) -> dict[str, str]:
    workbook = write_workbook(
        tmp_path / "book.xlsx",
        {"Inputs": {"A1": value}, "Outputs": {"B1": formula}},
    )
    document = bindings_document(
        series_entry(series_id, "Inputs!A1", layout="scalar", direction="input", dtype=dtype),
        series_entry(
            "result",
            "Outputs!B1",
            layout="scalar",
            direction="output",
            dtype=output_dtype,
            compute_name="compute_result",
        ),
        schema_version="1.19.0",
    )
    return generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_constraints({"Inputs!A1": constraint}),
    )


@pytest.mark.parametrize(
    ("kind", "expected"),
    [
        ("between", frozenset({"Annotated", "Between"})),
        ("real_between", frozenset({"Annotated", "RealBetween"})),
        ("literal", frozenset({"Literal"})),
        ("union", frozenset({"Annotated", "Literal", "RealBetween"})),
    ],
)
def test_scalar_domain_imports_only_where_the_annotation_is_spliced(
    tmp_path: Path, kind: str, expected: frozenset[str]
) -> None:
    """Scalar annotations are written into internals, validation, and model."""
    if kind == "between":
        modules = _scalar_pair(
            tmp_path,
            series_id="year",
            dtype="int",
            value=2000,
            output_dtype="int",
            formula="=Inputs!A1",
            constraint=Annotated[int, Between(1990, 2100)],
        )
    elif kind == "real_between":
        modules = _scalar_pair(
            tmp_path,
            series_id="share",
            dtype="float",
            value=0.25,
            output_dtype="float",
            formula="=Inputs!A1",
            constraint=Annotated[float, RealBetween(0, 1)],
        )
    elif kind == "literal":
        modules = _scalar_pair(
            tmp_path,
            series_id="flag",
            dtype="string",
            value="On",
            output_dtype="int",
            formula='=IF(Inputs!A1="On", 1, 0)',
            constraint=Literal["Off", "On"],
        )
    else:
        modules = _scalar_pair(
            tmp_path,
            series_id="sentinel",
            dtype="float",
            value="n.a.",
            output_dtype="float",
            formula="=IF(ISNUMBER(Inputs!A1), Inputs!A1, 0)",
            constraint=Literal["n.a."] | Annotated[float, RealBetween(-1.0, 1.0)],
        )
    spliced = {"internals.py", "validation.py", "model.py"}
    for name in _MODULES:
        found = _domain_imports(modules[name])
        assert found == (expected if name in spliced else frozenset())
    load_package(modules, tmp_path, name=f"scalar_{kind}")


def test_series_annotation_is_imported_by_data_and_validation_only(tmp_path: Path) -> None:
    """A public alias keeps the constructors out of modules that name `data.Choice`."""
    workbook = write_workbook(
        tmp_path / "book.xlsx",
        {
            "Inputs": {
                "A2": "grace",
                "A3": "maturity",
                "C2": "On",
                "C3": "Off",
                "D2": 1,
                "D3": 2,
                "I2": "=C2",
                "I3": "=C3",
                "J2": "=D2",
                "J3": "=D3",
            }
        },
    )
    document = bindings_document(
        _series("choice", "Inputs!C2:C3", direction="input", dtype="string"),
        _series("year", "Inputs!D2:D3", direction="input", dtype="int"),
        _series("choice_out", "Inputs!I2:I3", direction="output", dtype="string"),
        _series("year_out", "Inputs!J2:J3", direction="output", dtype="int"),
        schema_version="1.19.0",
    )
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_constraints(
            {
                "Inputs!C2": Literal["Off", "On"],
                "Inputs!C3": Literal["Off", "On"],
                "Inputs!D2": Annotated[int, Between(0, 5)],
                "Inputs!D3": Annotated[int, Between(0, 5)],
            }
        ),
    )
    expected = {
        "data.py": frozenset({"Annotated", "Between", "Literal"}),
        "validation.py": frozenset({"Annotated", "Between", "Literal"}),
        "internals.py": frozenset(),
        "model.py": frozenset(),
        "api.py": frozenset(),
    }
    for name, imports in expected.items():
        assert _domain_imports(modules[name]) == imports
    assert 'Literal["Off", "On"]' in modules["data.py"]
    assert "Between(" in modules["data.py"]
    assert 'Literal["Off", "On"]' in modules["validation.py"]
    assert "Between(" in modules["validation.py"]
    for name in ("internals.py", "model.py", "api.py"):
        assert 'Literal["Off", "On"]' not in modules[name]
        assert "Between(" not in modules[name]
    load_package(modules, tmp_path, name="series_domains")


def test_data_imports_span_only_for_an_emitted_call(tmp_path: Path) -> None:
    """A subset axis calls `span`; a key or cell that merely contains the text does not."""
    called = write_workbook(
        tmp_path / "called.xlsx",
        {
            "Data": {
                "B1": 2024,
                "C1": 2025,
                "D1": 2026,
                "B2": 1,
                "C2": 2,
                "D2": 3,
                "C4": 2025,
                "D4": 2026,
                "C5": 10,
                "D5": 11,
                "B7": "=SUM(B2:D2)+SUM(C5:D5)",
            }
        },
    )
    called_doc = bindings_document(
        series_entry(
            "prices_full",
            "Data!B2:D2",
            layout="series",
            direction="input",
            dtype="float",
            header_row=1,
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
        series_entry(
            "prices_short",
            "Data!C5:D5",
            layout="series",
            direction="input",
            dtype="float",
            header_row=4,
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
        series_entry(
            "total",
            "Data!B7",
            layout="scalar",
            direction="output",
            dtype="float",
            compute_name="compute_total",
        ),
        schema_version="1.19.0",
    )
    modules = generate_inverted(called, called_doc)
    assert "span(TIME_PERIOD_AXIS, 2025, 2026)" in modules["data.py"]
    assert "span" in _imported(modules["data.py"])
    assert "span(data.TIME_PERIOD_AXIS, 2025, 2026)" in modules["internals.py"]
    assert "span" in _imported(modules["internals.py"])
    for name in ("validation.py", "model.py", "api.py"):
        assert "span" not in _imported(modules[name])

    mentioned = _export_text(tmp_path, "span", "span(0, 1)", where="cell")
    assert "span(0, 1)" in mentioned["data.py"]
    assert "span" not in _imported(mentioned["data.py"])
    axis_dir = tmp_path / "axis"
    axis_dir.mkdir()
    keyed = _export_text(axis_dir, "notes", "span(0, 1)", where="axis")
    assert "span(0, 1)" in keyed["data.py"]
    assert "span" not in _imported(keyed["data.py"])
