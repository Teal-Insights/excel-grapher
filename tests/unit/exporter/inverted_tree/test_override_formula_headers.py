"""Override inputs keep formula headers on the export snapshot.

`input.mode: override` displaces template values. Callers type over the
formula; the header text stays the snapshot. A leaf input may still rekey a
formula header from the case workbook. An override input must not become
that runtime labeller, or every other series on the same snapshot keys
inherits the axis and cross-field matches fail.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    make_catalog,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_formula_header_string_remap import _plant_strings


def _header_series(
    series_id: str,
    keys: tuple[str, ...],
    *,
    column: str,
    mode: str | None,
) -> BoundSeries:
    cells = tuple(f"Terms!{column}{index + 1}" for index in range(len(keys)))
    domain = tuple(KeyPoint((("VARIANT", key),)) for key in keys)
    raw: dict[str, object] = {"input": {}} if mode is None else {"input": {"mode": mode}}
    return BoundSeries(
        series_id=series_id,
        layout="series",
        direction="input",
        cells=cells,
        key_fields=("VARIANT",),
        dtype="float",
        compute_name=None,
        raw=raw,
        domain=domain,
        statements=(Statement(series_id, series_id, None, 0, len(cells), cells, domain),),
        formula_label_addresses=(("VARIANT", ("Terms!H7", "Terms!I7")),),
    )


def test_header_label_index_skips_override_inputs() -> None:
    """A constructed override input is not a runtime labeller for its headers."""
    override = _header_series("terms", ("Default", "User defined"), column="A", mode="override")
    leaf = _header_series("amounts", ("Default", "User defined"), column="B", mode=None)
    catalog = make_catalog(
        {"terms": override, "amounts": leaf},
        ("terms", "amounts"),
        {
            cell: series_id
            for series_id, series in (("terms", override), ("amounts", leaf))
            for cell in series.cells
        },
    )
    assert catalog.runtime_labeller("VARIANT", ("Default", "User defined")) is leaf
    assert catalog.header_labeller_ids() == {"VARIANT": "amounts"}

    override_catalog = make_catalog(
        {"terms": override},
        ("terms",),
        {cell: "terms" for cell in override.cells},
    )
    assert override_catalog.runtime_labeller("VARIANT", ("Default", "User defined")) is None
    assert override_catalog.header_labeller_ids() == {}


def _row_bindings() -> dict:
    leaf = series_entry(
        "amounts",
        "Inputs!B2:B3",
        layout="series",
        direction="input",
        label_column="A",
        key_concept="NAME",
        key_read="string",
    )
    override = series_entry(
        "rates",
        "Inputs!D2:D3",
        layout="series",
        direction="input",
        label_column="C",
        key_concept="NAME",
        key_read="string",
    )
    override["input"]["mode"] = "override"
    document = bindings_document(leaf, override, schema_version="1.22.0")
    document["concept_scheme"]["concepts"].append({"id": "NAME", "dtype": "string"})
    return document


def test_override_row_labels_are_not_formula_label_addresses(tmp_path: Path) -> None:
    """Formula row labels on an override input stay off the header-label index."""
    path = tmp_path / "rows.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A2": '="Alpha"',
                "A3": '="Beta"',
                "B2": 1,
                "B3": 2,
                "C2": '="Gamma"',
                "C3": '="Delta"',
                "D2": "=10",
                "D3": "=20",
            }
        },
    )
    _plant_strings(path, {"A2": "Alpha", "A3": "Beta", "C2": "Gamma", "C3": "Delta"})
    document = _row_bindings()
    catalog, _deps, _graph = inverted_graph_parts(path, document)
    assert catalog.get("amounts").formula_label_addresses == (("NAME", ("Inputs!A2", "Inputs!A3")),)
    assert catalog.get("rates").formula_label_addresses == ()
    assert catalog.runtime_labeller("NAME", ("Alpha", "Beta")) is catalog.get("amounts")
    assert catalog.runtime_labeller("NAME", ("Gamma", "Delta")) is None


def _dimension(axis_id: str, bind: dict) -> dict:
    return {
        "id": axis_id,
        "concept": axis_id,
        "role": "key",
        "scope": "cell",
        "bind": bind,
    }


def _column_header(axis_id: str) -> dict:
    return _dimension(
        axis_id,
        {"kind": "column_header", "header_row": 7, "read": "string", "normalize": "strip"},
    )


def _row_label(axis_id: str) -> dict:
    return _dimension(
        axis_id,
        {"kind": "row_label", "label_column": "G", "read": "string", "normalize": "strip"},
    )


def _measure() -> dict:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _terms_bindings() -> dict:
    return {
        "schema_version": "1.22.0",
        "concept_scheme": {
            "id": "terms",
            "concepts": [
                {"id": "VARIANT", "dtype": "string"},
                {"id": "INDICATOR", "dtype": "string"},
                {"id": "INSTRUMENT", "dtype": "string"},
                {"id": "OBS_VALUE", "dtype": "number"},
            ],
        },
        "series": [
            {
                "id": "terms",
                "sheet": "Terms",
                "data_range": "Terms!H14:I14",
                "layout": "matrix",
                "input": {"mode": "override"},
                "key": ["INDICATOR", "VARIANT"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [_row_label("INDICATOR"), _column_header("VARIANT")],
                },
            },
            {
                "id": "shares",
                "sheet": "Terms",
                "data_range": "Terms!H9:I10",
                "layout": "matrix",
                "internal": {},
                "key": ["INSTRUMENT", "VARIANT"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [_row_label("INSTRUMENT"), _column_header("VARIANT")],
                },
            },
            {
                "id": "fallback",
                "sheet": "Terms",
                "data_range": "Terms!H11",
                "layout": "scalar",
                "internal": {},
                "key": [],
                "structure": {"measure": _measure(), "dimensions": []},
            },
            {
                "id": "domestic_st",
                "sheet": "Terms",
                "data_range": "Terms!I11:J11",
                "layout": "series",
                "output": {"compute": {"name": "compute_domestic_st"}},
                "key": ["INDICATOR"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [_column_header("INDICATOR")],
                },
            },
        ],
    }


def _write_terms(path: Path) -> None:
    write_workbook(
        path,
        {
            "Terms": {
                "H7": '="Default"',
                "I7": '="User defined"',
                "J7": "Effective",
                "G9": "External",
                "H9": 0.6,
                "I9": "=H9",
                "G10": "Domestic",
                "H10": 0.3,
                "I10": "=H10",
                "H11": "=0.1",
                "I11": "=1-I9-I10",
                "J11": "=IF(ISNUMBER(I11),I11,H11)",
                "G14": "Rate",
                "H14": "=0.05",
                "I14": "=H14",
            }
        },
    )
    _plant_strings(path, {"H7": "Default", "I7": "User defined"})


def _assert_matches_evaluator(model: object, series: BoundSeries, graph: object) -> None:
    expected = FormulaEvaluator(graph).evaluate(list(series.cells))
    got = getattr(model, series.series_id)
    if series.is_scalar:
        assert got == pytest.approx(expected[series.cells[0]])
        return
    for (_coord, value), cell in zip(got.items(), series.cells, strict=True):
        assert value == pytest.approx(expected[cell])


def test_override_formula_headers_do_not_label_other_series(tmp_path: Path) -> None:
    """A formula header on an override input stays a static snapshot key.

    `domestic_st` keys column I as the static indicator `User defined` and
    reads `shares` on that same header. Export must not treat the override
    series as the runtime labeller for that header.
    """
    path = tmp_path / "terms.xlsx"
    _write_terms(path)
    document = _terms_bindings()
    modules = generate_inverted(path, document)
    assert "LABEL_CELLS" not in modules["data.py"]
    catalog, _deps, graph = inverted_graph_parts(path, document)
    assert catalog.get("terms").formula_label_addresses == ()
    assert catalog.runtime_labeller("VARIANT", ("Default", "User defined")) is None
    assert "VARIANT" not in catalog.header_labeller_ids()
    pkg = load_package(modules, tmp_path, name="override_formula_headers")
    model = pkg.Model.from_defaults()
    loaded = pkg.Model.from_workbook(path)
    for series_id in ("terms", "shares", "fallback", "domestic_st"):
        series = catalog.get(series_id)
        _assert_matches_evaluator(model, series, graph)
        _assert_matches_evaluator(loaded, series, graph)
    assert tuple(model.domestic_st.domain.axes[0].keys) == ("User defined", "Effective")

    case = tmp_path / "case.xlsx"
    _write_terms(case)
    _plant_strings(case, {"H7": "Renamed", "I7": "Also renamed"})
    rebound = pkg.Model.from_workbook(case)
    assert tuple(axis.name for axis in rebound.terms.domain.axes) == ("INDICATOR", "VARIANT")
    assert tuple(rebound.terms.domain.axes[1].keys) == ("Default", "User defined")
    assert tuple(rebound.domestic_st.domain.axes[0].keys) == ("User defined", "Effective")
