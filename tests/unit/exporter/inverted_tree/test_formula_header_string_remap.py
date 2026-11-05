"""A formula-labelled string axis follows another vocabulary by position.

Snapshot spellings are the keys a case workbook replaces. A lockstep copy
onto a different producer vocabulary is the producer key at the host's
labeller position, not a dict of those spellings. A wider labeller, a
matrix labeller, and a shared integer axis keep the positions they already
had: the labelled axis is selected by name, and a period index is a
different scale from the strings.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _plant_strings(path: Path, labels: dict[str, str]) -> None:
    """Keep each formula and set its saved string value."""
    sheet = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(sheet).decode()
        others = {name: archive.read(name) for name in archive.namelist() if name != sheet}
    for coordinate, cached in labels.items():
        marker = f'<c r="{coordinate}"'
        start = xml.find(marker)
        end = xml.find("</c>", start)
        if start < 0 or end < 0:
            raise AssertionError(f"{coordinate} formula cell not found")
        cell = xml[start:end]
        if "<f>" not in cell and "<f " not in cell:
            raise AssertionError(f"{coordinate} is not a formula cell: {cell}")
        if 't="' not in cell:
            cell = cell.replace(marker, f'{marker} t="str"', 1)
        if "<v />" in cell:
            cell = cell.replace("<v />", f"<v>{cached}</v>", 1)
        elif "<v/>" in cell:
            cell = cell.replace("<v/>", f"<v>{cached}</v>", 1)
        elif "<v>" in cell:
            close = cell.find("</v>")
            open_end = cell.find(">", cell.find("<v"))
            cell = cell[: open_end + 1] + cached + cell[close:]
        else:
            cell += f"<v>{cached}</v>"
        xml = xml[:start] + cell + xml[end:]
    with zipfile.ZipFile(path, "w") as archive:
        for name, payload in others.items():
            archive.writestr(name, payload)
        archive.writestr(sheet, xml)


def _document(*series: dict[str, Any]) -> dict[str, Any]:
    document = bindings_document(*series, schema_version="1.22.0")
    known = {item["id"] for item in document["concept_scheme"]["concepts"]}
    for concept in ("COMMODITY", "COUNTRY", "INSTRUMENT"):
        if concept not in known:
            document["concept_scheme"]["concepts"].append({"id": concept, "dtype": "string"})
    return document


def _row(
    series_id: str,
    data_range: str,
    label_column: str,
    *,
    direction: str,
    key: str,
    dtype: str = "float",
) -> dict[str, Any]:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction=direction,
        dtype=dtype,
        label_column=label_column,
        key_concept=key,
        key_read="string",
    )


def _column(
    series_id: str,
    data_range: str,
    header_row: int,
    *,
    direction: str,
    key: str,
    dtype: str = "float",
    read: str = "string",
) -> dict[str, Any]:
    return series_entry(
        series_id,
        data_range,
        layout="series",
        direction=direction,
        dtype=dtype,
        header_row=header_row,
        key_concept=key,
        key_read=read,
    )


def _dim(concept: str, bind: dict[str, Any]) -> dict[str, Any]:
    return {"id": concept, "concept": concept, "role": "key", "scope": "cell", "bind": bind}


def _grid(
    series_id: str,
    data_range: str,
    dims: list[dict[str, Any]],
    key: list[str],
    *,
    direction: str,
    dtype: str = "float",
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": data_range.split("!", 1)[0],
        "data_range": data_range,
        "layout": "series",
        "key": key,
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": dtype,
                "bind": {"kind": "data_cell", "read": dtype if dtype != "number" else "float"},
            },
            "dimensions": dims,
        },
    }
    if direction == "input":
        entry["input"] = {}
    else:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    return entry


def _compute_body(source: str, series_id: str) -> str:
    for marker in (f"def compute_{series_id}", f"def {series_id}("):
        start = source.find(marker)
        if start < 0:
            continue
        rest = source[start:]
        nxt = rest.find("\ndef ", 1)
        return rest if nxt < 0 else rest[:nxt]
    raise AssertionError(source)


def _values_by_cell(model: object, series: BoundSeries, graph: object) -> list[object]:
    expected = FormulaEvaluator(graph).evaluate(list(series.cells))
    got = getattr(model, series.series_id)
    found = []
    for (_coord, value), cell in zip(got.items(), series.cells, strict=True):
        exp = expected[cell]
        if isinstance(exp, float) and isinstance(value, float):
            assert value == pytest.approx(exp)
        else:
            assert value == exp
        found.append(value)
    return found


def _load(tmp_path: Path, path: Path, document: dict[str, Any], name: str):
    modules = generate_inverted(path, document)
    pkg = load_package(modules, tmp_path, name=name)
    catalog, _deps, graph = inverted_graph_parts(path, document)
    return modules, pkg, catalog, graph


def test_formula_labelled_host_follows_literal_producer_by_position(tmp_path: Path) -> None:
    """`Base`/`Grain` copy `Copper`/`Wheat` without a snapshot dict.

    The two vocabularies share the axis name and sort into one interned
    constant (`Base`, `Copper`, `Grain`, `Wheat`). The follow indexes the
    producer keys, so that union does not move `Copper`.
    """
    path = tmp_path / "commodities.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A2": '="Base"',
                "A3": '="Grain"',
                "B2": "Yes",
                "B3": "No",
                "C2": "=E2",
                "C3": "=E3",
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(path, {"A2": "Base", "A3": "Grain"})
    document = _document(
        _row("flags", "Inputs!B2:B3", "A", direction="input", key="COMMODITY", dtype="string"),
        _row("prices", "Inputs!E2:E3", "D", direction="input", key="COMMODITY"),
        _row("shock", "Inputs!C2:C3", "A", direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "formula_string_remap")
    body = _compute_body(modules["internals.py"], "shock")
    assert "('Copper', 'Wheat')" in body
    assert ".position(" in body
    assert "Base" not in body
    assert "Grain" not in body
    assert "COMMODITY_TO_" not in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)
    assert model.shock[("Base",)] == pytest.approx(10)
    assert model.shock[("Grain",)] == pytest.approx(20)

    case = tmp_path / "translated.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A2": '="Base"',
                "A3": '="Grain"',
                "B2": "Yes",
                "B3": "No",
                "C2": "=E2",
                "C3": "=E3",
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(case, {"A2": "Metal", "A3": "Cereal"})
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(case, document)
    shifted = pkg.Model.from_workbook(case)
    _values_by_cell(shifted, shifted_catalog.get("shock"), shifted_graph)
    assert tuple(shifted.flags.domain.axes[0].keys) == ("Metal", "Cereal")
    assert tuple(shifted.prices.domain.axes[0].keys) == ("Copper", "Wheat")
    assert shifted.shock[("Metal",)] == pytest.approx(10)
    assert shifted.shock[("Cereal",)] == pytest.approx(20)


def test_host_window_of_a_wider_labeller_shifts_by_that_offset(tmp_path: Path) -> None:
    """`shock` starts on the labeller's second key and reads the producer from zero."""
    path = tmp_path / "window.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A2": '="Hist"',
                "A3": '="Base"',
                "A4": '="Grain"',
                "B2": 1,
                "B3": 2,
                "B4": 3,
                "C3": "=E3",
                "C4": "=E4",
                "D3": "Copper",
                "D4": "Wheat",
                "E3": 10,
                "E4": 20,
            }
        },
    )
    _plant_strings(path, {"A2": "Hist", "A3": "Base", "A4": "Grain"})
    document = _document(
        _row("flags", "Inputs!B2:B4", "A", direction="input", key="COMMODITY", dtype="float"),
        _row("prices", "Inputs!E3:E4", "D", direction="input", key="COMMODITY"),
        _row("shock", "Inputs!C3:C4", "A", direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "window_string_remap")
    body = _compute_body(modules["internals.py"], "shock")
    assert "('Copper', 'Wheat')" in body
    assert "- 1" in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)
    assert model.shock[("Base",)] == pytest.approx(10)
    assert model.shock[("Grain",)] == pytest.approx(20)

    case = tmp_path / "window_case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A2": '="Hist"',
                "A3": '="Base"',
                "A4": '="Grain"',
                "B2": 1,
                "B3": 2,
                "B4": 3,
                "C3": "=E3",
                "C4": "=E4",
                "D3": "Copper",
                "D4": "Wheat",
                "E3": 10,
                "E4": 20,
            }
        },
    )
    _plant_strings(case, {"A2": "Old", "A3": "Metal", "A4": "Cereal"})
    shifted = pkg.Model.from_workbook(case)
    assert tuple(shifted.shock.domain.axes[0].keys) == ("Metal", "Cereal")
    assert shifted.shock[("Metal",)] == pytest.approx(10)
    assert shifted.shock[("Cereal",)] == pytest.approx(20)


def test_same_labeller_spelling_passes_the_host_key_through(tmp_path: Path) -> None:
    """Matching snapshot text on one labeller stays the host variable."""
    path = tmp_path / "same.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A2": '="Copper"',
                "A3": '="Wheat"',
                "B2": "Yes",
                "B3": "No",
                "C2": "=E2",
                "C3": "=E3",
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(path, {"A2": "Copper", "A3": "Wheat"})
    document = _document(
        _row("flags", "Inputs!B2:B3", "A", direction="input", key="COMMODITY", dtype="string"),
        _row("prices", "Inputs!E2:E3", "D", direction="input", key="COMMODITY"),
        _row("shock", "Inputs!C2:C3", "A", direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "same_labeller_spelling")
    body = _compute_body(modules["internals.py"], "shock")
    assert "prices[commodity]" in body
    assert "('Copper', 'Wheat')" not in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)

    case = tmp_path / "same_case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A2": '="Copper"',
                "A3": '="Wheat"',
                "B2": "Yes",
                "B3": "No",
                "C2": "=E2",
                "C3": "=E3",
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(case, {"A2": "Base", "A3": "Grain"})
    shifted = pkg.Model.from_workbook(case)
    assert tuple(shifted.prices.domain.axes[0].keys) == ("Base", "Grain")
    assert shifted.shock[("Base",)] == pytest.approx(10)
    assert shifted.shock[("Grain",)] == pytest.approx(20)


def test_same_labeller_row_shift_indexes_live_keys(tmp_path: Path) -> None:
    """A one-row lag on the labelling series follows the loaded labels."""
    path = tmp_path / "shift.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A2": '="Base"',
                "A3": '="Grain"',
                "A4": '="Ore"',
                "B2": 10,
                "B3": 20,
                "B4": 30,
                "C3": "=B2",
                "C4": "=B3",
            }
        },
    )
    _plant_strings(path, {"A2": "Base", "A3": "Grain", "A4": "Ore"})
    document = _document(
        _row("source", "Inputs!B2:B4", "A", direction="input", key="COMMODITY"),
        _row("shock", "Inputs!C3:C4", "A", direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "same_labeller_shift")
    body = _compute_body(modules["internals.py"], "shock")
    assert "source.domain.axis('COMMODITY').keys" in body
    assert "- 1" in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)
    assert model.shock[("Grain",)] == pytest.approx(10)
    assert model.shock[("Ore",)] == pytest.approx(20)

    case = tmp_path / "shift_case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A2": '="Base"',
                "A3": '="Grain"',
                "A4": '="Ore"',
                "B2": 10,
                "B3": 20,
                "B4": 30,
                "C3": "=B2",
                "C4": "=B3",
            }
        },
    )
    _plant_strings(case, {"A2": "X", "A3": "Y", "A4": "Z"})
    shifted = pkg.Model.from_workbook(case)
    assert tuple(shifted.source.domain.axes[0].keys) == ("X", "Y", "Z")
    assert tuple(shifted.shock.domain.axes[0].keys) == ("Y", "Z")
    assert shifted.shock[("Y",)] == pytest.approx(10)
    assert shifted.shock[("Z",)] == pytest.approx(20)


def test_matching_text_on_another_axis_does_not_pass_through(tmp_path: Path) -> None:
    """The same snapshot words on `COUNTRY` are still that axis's own keys."""
    path = tmp_path / "cross.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A2": '="Copper"',
                "A3": '="Wheat"',
                "B2": 1,
                "B3": 1,
                "C2": "=E2",
                "C3": "=E3",
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(path, {"A2": "Copper", "A3": "Wheat"})
    document = _document(
        _row("flags", "Inputs!B2:B3", "A", direction="input", key="COMMODITY"),
        _row("prices", "Inputs!E2:E3", "D", direction="input", key="COUNTRY"),
        _row("shock", "Inputs!C2:C3", "A", direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "cross_axis_identity")
    body = _compute_body(modules["internals.py"], "shock")
    assert "prices[commodity]" not in body
    assert "('Copper', 'Wheat')" in body
    assert ".position(" in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)

    case = tmp_path / "cross_case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A2": '="Copper"',
                "A3": '="Wheat"',
                "B2": 1,
                "B3": 1,
                "C2": "=E2",
                "C3": "=E3",
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(case, {"A2": "Base", "A3": "Grain"})
    shifted = pkg.Model.from_workbook(case)
    assert tuple(shifted.prices.domain.axes[0].keys) == ("Copper", "Wheat")
    assert shifted.shock[("Base",)] == pytest.approx(10)
    assert shifted.shock[("Grain",)] == pytest.approx(20)


def test_two_formula_axes_follow_each_other_by_position(tmp_path: Path) -> None:
    """Host and producer labels can both change and still align by row."""
    path = tmp_path / "both.xlsx"

    def sheet() -> dict[str, object]:
        return {
            "A2": '="Copper"',
            "A3": '="Wheat"',
            "B2": 1,
            "B3": 1,
            "C2": "=E2",
            "C3": "=E3",
            "D2": '="Copper"',
            "D3": '="Wheat"',
            "E2": 10,
            "E3": 20,
        }

    write_workbook(path, {"Inputs": sheet()})
    _plant_strings(path, {"A2": "Copper", "A3": "Wheat", "D2": "Copper", "D3": "Wheat"})
    document = _document(
        _row("flags", "Inputs!B2:B3", "A", direction="input", key="COMMODITY"),
        _row("prices", "Inputs!E2:E3", "D", direction="input", key="COUNTRY"),
        _row("shock", "Inputs!C2:C3", "A", direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "two_formula_axes")
    body = _compute_body(modules["internals.py"], "shock")
    assert "prices.domain.axis('COUNTRY').keys" in body
    assert ".position(" in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)

    case = tmp_path / "both_case.xlsx"
    write_workbook(case, {"Inputs": sheet()})
    _plant_strings(case, {"A2": "Base", "A3": "Grain", "D2": "Red", "D3": "Blue"})
    shifted = pkg.Model.from_workbook(case)
    assert tuple(shifted.flags.domain.axes[0].keys) == ("Base", "Grain")
    assert tuple(shifted.prices.domain.axes[0].keys) == ("Red", "Blue")
    assert shifted.shock[("Base",)] == pytest.approx(10)
    assert shifted.shock[("Grain",)] == pytest.approx(20)
    assert shifted.prices[("Red",)] == pytest.approx(10)


def test_if_copies_the_producer_row_when_the_flag_says_yes(tmp_path: Path) -> None:
    """One expression gates the positional copy; it does not branch on the label."""
    path = tmp_path / "gated.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A2": '="Base"',
                "A3": '="Grain"',
                "B2": "Yes",
                "B3": "No",
                "C2": '=IF(B2="Yes",E2,"")',
                "C3": '=IF(B3="Yes",E3,"")',
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(path, {"A2": "Base", "A3": "Grain"})
    document = _document(
        _row("flags", "Inputs!B2:B3", "A", direction="input", key="COMMODITY", dtype="string"),
        _row("prices", "Inputs!E2:E3", "D", direction="input", key="COMMODITY"),
        _row("shock", "Inputs!C2:C3", "A", direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "gated_string_remap")
    body = _compute_body(modules["internals.py"], "shock")
    assert "('Copper', 'Wheat')" in body
    assert ".position(" in body
    assert "Base" not in body
    assert "commodity ==" not in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)
    assert model.shock[("Base",)] == pytest.approx(10)
    assert model.shock[("Grain",)] == ""

    case = tmp_path / "gated_case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "A2": '="Base"',
                "A3": '="Grain"',
                "B2": "Yes",
                "B3": "No",
                "C2": '=IF(B2="Yes",E2,"")',
                "C3": '=IF(B3="Yes",E3,"")',
                "D2": "Copper",
                "D3": "Wheat",
                "E2": 10,
                "E3": 20,
            }
        },
    )
    _plant_strings(case, {"A2": "Metal", "A3": "Cereal"})
    shifted = pkg.Model.from_workbook(case)
    assert shifted.shock[("Metal",)] == pytest.approx(10)
    assert shifted.shock[("Cereal",)] == ""


def test_matrix_labeller_selects_the_string_axis_by_name(tmp_path: Path) -> None:
    """The widest formula-header input is a matrix; its first axis is not the label."""
    path = tmp_path / "matrix.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "B1": '="Base"',
                "C1": '="Grain"',
                "A2": "IMF",
                "B2": 1,
                "C2": 2,
                "A3": "IDA",
                "B3": 3,
                "C3": 4,
                "B5": "Copper",
                "C5": "Wheat",
                "B6": 10,
                "C6": 20,
                "B7": "=B6",
                "C7": "=C6",
            }
        },
    )
    _plant_strings(path, {"B1": "Base", "C1": "Grain"})
    flags = {
        "id": "flags",
        "sheet": "Inputs",
        "data_range": "Inputs!B2:C3",
        "layout": "series",
        "input": {},
        "key": ["INSTRUMENT", "COMMODITY"],
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                _dim(
                    "INSTRUMENT",
                    {"kind": "value_map", "values": {"IMF": 2, "IDA": 3}, "read": "string"},
                ),
                _dim(
                    "COMMODITY",
                    {"kind": "column_header", "header_row": 1, "read": "string"},
                ),
            ],
        },
    }
    document = _document(
        flags,
        _column("prices", "Inputs!B6:C6", 5, direction="input", key="COMMODITY"),
        _column("shock", "Inputs!B7:C7", 1, direction="output", key="COMMODITY"),
    )
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "matrix_string_remap")
    body = _compute_body(modules["internals.py"], "shock")
    assert "('Copper', 'Wheat')" in body
    assert ".position(" in body
    assert "axis('INSTRUMENT')" not in body
    service = catalog.get("flags")
    assert tuple(axis.name for axis in service.tensor_domain.axes) == ("INSTRUMENT", "COMMODITY")
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)
    assert model.shock[("Base",)] == pytest.approx(10)
    assert model.shock[("Grain",)] == pytest.approx(20)

    case = tmp_path / "matrix_case.xlsx"
    write_workbook(
        case,
        {
            "Inputs": {
                "B1": '="Base"',
                "C1": '="Grain"',
                "A2": "IMF",
                "B2": 1,
                "C2": 2,
                "A3": "IDA",
                "B3": 3,
                "C3": 4,
                "B5": "Copper",
                "C5": "Wheat",
                "B6": 10,
                "C6": 20,
                "B7": "=B6",
                "C7": "=C6",
            }
        },
    )
    _plant_strings(case, {"B1": "Metal", "C1": "Cereal"})
    shifted = pkg.Model.from_workbook(case)
    assert tuple(shifted.flags.domain.axes[1].keys) == ("Metal", "Cereal")
    assert shifted.shock[("Metal",)] == pytest.approx(10)
    assert shifted.shock[("Cereal",)] == pytest.approx(20)


def _year_block(*, swapped: bool) -> dict[str, object]:
    host = {
        "B2": "=E3" if swapped else "=E2",
        "C2": "=F3" if swapped else "=F2",
        "B3": "=E2" if swapped else "=E3",
        "C3": "=F2" if swapped else "=F3",
    }
    return {
        "B1": 2024,
        "C1": 2025,
        "A2": '="Base"',
        "A3": '="Grain"',
        "H2": 1,
        "H3": 1,
        "D2": "Red",
        "D3": "Blue",
        "E1": 2024,
        "F1": 2025,
        "E2": 10,
        "F2": 11,
        "E3": 20,
        "F3": 21,
        **host,
    }


def _year_document() -> dict[str, Any]:
    row = {"kind": "row_label", "label_column": "A", "read": "string"}
    year = {"kind": "column_header", "header_row": 1, "read": "int"}
    return _document(
        _row("flags", "Inputs!H2:H3", "A", direction="input", key="COMMODITY"),
        _grid(
            "prices",
            "Inputs!E2:F3",
            [
                _dim("COUNTRY", {"kind": "row_label", "label_column": "D", "read": "string"}),
                _dim("TIME_PERIOD", year),
            ],
            ["COUNTRY", "TIME_PERIOD"],
            direction="input",
        ),
        _grid(
            "shock",
            "Inputs!B2:C3",
            [_dim("COMMODITY", row), _dim("TIME_PERIOD", year)],
            ["COMMODITY", "TIME_PERIOD"],
            direction="output",
        ),
    )


def test_string_remap_keeps_the_shared_year_as_the_host_variable(tmp_path: Path) -> None:
    """Commodity follows by position; `TIME_PERIOD` is the same literal scale."""
    path = tmp_path / "grid.xlsx"
    write_workbook(path, {"Inputs": _year_block(swapped=False)})
    _plant_strings(path, {"A2": "Base", "A3": "Grain"})
    document = _year_document()
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "grid_string_remap")
    body = _compute_body(modules["internals.py"], "shock")
    assert "('Red', 'Blue')" in body
    assert "time_period" in body
    assert ".position(" in body
    assert "Base" not in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)
    assert model.shock[("Base", 2024)] == pytest.approx(10)
    assert model.shock[("Base", 2025)] == pytest.approx(11)
    assert model.shock[("Grain", 2024)] == pytest.approx(20)
    assert model.shock[("Grain", 2025)] == pytest.approx(21)

    case = tmp_path / "grid_case.xlsx"
    write_workbook(case, {"Inputs": _year_block(swapped=False)})
    _plant_strings(case, {"A2": "Metal", "A3": "Cereal"})
    shifted = pkg.Model.from_workbook(case)
    assert shifted.shock[("Metal", 2024)] == pytest.approx(10)
    assert shifted.shock[("Cereal", 2025)] == pytest.approx(21)


def test_swapped_rows_use_a_position_dict(tmp_path: Path) -> None:
    """A commodity swap is not one offset, and it still is not a snapshot dict."""
    path = tmp_path / "swap.xlsx"
    write_workbook(path, {"Inputs": _year_block(swapped=True)})
    _plant_strings(path, {"A2": "Base", "A3": "Grain"})
    document = _year_document()
    modules, pkg, catalog, graph = _load(tmp_path, path, document, "swap_string_remap")
    body = _compute_body(modules["internals.py"], "shock")
    assert "{0: 1, 1: 0}" in body
    assert "('Red', 'Blue')" in body
    assert "Base" not in body
    model = pkg.Model.from_workbook(path)
    _values_by_cell(model, catalog.get("shock"), graph)
    assert model.shock[("Base", 2024)] == pytest.approx(20)
    assert model.shock[("Base", 2025)] == pytest.approx(21)
    assert model.shock[("Grain", 2024)] == pytest.approx(10)
    assert model.shock[("Grain", 2025)] == pytest.approx(11)

    case = tmp_path / "swap_case.xlsx"
    write_workbook(case, {"Inputs": _year_block(swapped=True)})
    _plant_strings(case, {"A2": "Metal", "A3": "Cereal"})
    shifted = pkg.Model.from_workbook(case)
    assert shifted.shock[("Metal", 2024)] == pytest.approx(20)
    assert shifted.shock[("Cereal", 2025)] == pytest.approx(11)
