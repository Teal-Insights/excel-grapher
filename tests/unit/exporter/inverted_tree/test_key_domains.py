"""Issue 676 — key domains in data.py and __key__/__domain__ on compute_*."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.domains import (
    collect_field_domains,
    series_domain_points,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _years_workbook(tmp_path: Path, years: tuple[int, ...], *, extra_outputs: int = 1) -> Path:
    cells: dict[str, object] = {}
    for index, year in enumerate(years):
        col = chr(ord("B") + index)
        cells[f"{col}1"] = year
        cells[f"{col}2"] = float(index + 1)
        cells[f"{col}3"] = f"=Inputs!{col}2"
        for extra in range(1, extra_outputs):
            cells[f"{col}{3 + extra}"] = f"=Inputs!{col}2"
    return write_workbook(tmp_path / f"years_{len(years)}_{extra_outputs}.xlsx", {"Inputs": cells})


def _years_bindings(years: tuple[int, ...], *, extra_outputs: int = 1) -> dict[str, Any]:
    last = chr(ord("B") + len(years) - 1)
    entries = [
        series_entry(
            "growth",
            f"Inputs!B2:{last}2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "path",
            f"Inputs!B3:{last}3",
            layout="series",
            direction="output",
            header_row=1,
        ),
    ]
    for extra in range(1, extra_outputs):
        row = 3 + extra
        entries.append(
            series_entry(
                f"path_{extra}",
                f"Inputs!B{row}:{last}{row}",
                layout="series",
                direction="output",
                header_row=1,
            )
        )
    return bindings_document(*entries)


def _matrix_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "key_domain_matrix.xlsx",
        {
            "Engine": {
                "B1": 2020,
                "C1": 2021,
                "A2": "France",
                "B2": 100.0,
                "C2": 110.0,
                "A3": "Kenya",
                "B3": 50.0,
                "C3": 55.0,
                "B4": 2020,
                "C4": 2021,
                "A5": "France",
                "B5": 10.0,
                "C5": 11.0,
                "A6": "Kenya",
                "B6": 5.0,
                "C6": 6.0,
                "B7": 2020,
                "C7": 2021,
                "A8": "France",
                "B8": "=B5/B2",
                "C8": "=C5/C2",
                "A9": "Kenya",
                "B9": "=B6/B3",
                "C9": "=C6/C3",
            },
        },
    )


def _matrix_entry(series_id: str, data_range: str, *, header_row: int, direction: str) -> dict:
    sheet = data_range.split("!", 1)[0]
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": sheet,
        "data_range": data_range,
        "layout": "matrix",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "REF_AREA",
                    "concept": "REF_AREA",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
                },
            ],
        },
        "key": ["REF_AREA", "TIME_PERIOD"],
    }
    if direction == "output":
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    elif direction == "internal":
        entry["internal"] = {}
    else:
        entry["constant"] = {}
    return entry


def _matrix_bindings() -> dict:
    return bindings_document(
        _matrix_entry("gdp", "Engine!B2:C3", header_row=1, direction="constant"),
        _matrix_entry("revenue", "Engine!B5:C6", header_row=4, direction="constant"),
        _matrix_entry("ratio", "Engine!B8:C9", header_row=7, direction="output"),
    )


def _late_start_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "late_start.xlsx",
        {
            "Engine": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "A2": 1.0,
                "B2": 2.0,
                "C2": 3.0,
                "B3": "=B2",
                "C3": "=C2",
            },
        },
    )


def _late_start_bindings() -> dict:
    return bindings_document(
        series_entry(
            "values",
            "Engine!A2:C2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "adj",
            "Engine!B3:C3",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def test_collect_field_domains_is_first_seen_catalog_order(tmp_path: Path) -> None:
    years = (2008, 2009, 2010)
    workbook = _years_workbook(tmp_path, years)
    catalog = build_catalog(validate_bindings_document(_years_bindings(years)), workbook=workbook)
    domains = collect_field_domains(catalog)
    assert list(domains) == ["TIME_PERIOD"]
    assert domains["TIME_PERIOD"] == years
    assert series_domain_points(catalog.get("path")) == years


def test_data_module_emits_one_tuple_per_distinct_field(tmp_path: Path) -> None:
    years = (2008, 2009, 2010)
    small = generate_inverted(_years_workbook(tmp_path, years), _years_bindings(years))
    large = generate_inverted(
        _years_workbook(tmp_path, years, extra_outputs=6),
        _years_bindings(years, extra_outputs=6),
    )
    assert small["data.py"].count("TIME_PERIOD_DOMAIN") == large["data.py"].count(
        "TIME_PERIOD_DOMAIN"
    )
    assert small["data.py"].count("TIME_PERIOD_DOMAIN:") == 1
    assert large["data.py"].count("TIME_PERIOD_DOMAIN:") == 1
    assert "TIME_PERIOD_DOMAIN: tuple[int, ...] = (2008, 2009, 2010)" in small["data.py"]
    assert "TIME_PERIOD_DOMAIN: tuple[int, ...] = (2008, 2009, 2010)" in large["data.py"]


def test_compute_key_and_domain_match_result_length(tmp_path: Path) -> None:
    years = (2008, 2009, 2010)
    workbook = _years_workbook(tmp_path, years)
    pkg = load_package(
        generate_inverted(workbook, _years_bindings(years)), tmp_path, name="key_years"
    )
    assert pkg.compute_path.__key__ == ("TIME_PERIOD",)
    assert pkg.compute_path.__domain__ == years
    assert pkg.compute_path.__domain__ is pkg.data.TIME_PERIOD_DOMAIN
    result = pkg.compute_path(growth=pkg.data.GROWTH_DEFAULT)
    assert len(pkg.compute_path.__domain__) == len(result)
    assert result[pkg.data.TIME_PERIOD_DOMAIN.index(2009)] == pytest.approx(2.0)
    assert pkg.internals.path.__key__ == ("TIME_PERIOD",)
    assert pkg.internals.path.__domain__ == years


def test_scalar_compute_publishes_empty_key_and_unit_domain(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "scalar_key.xlsx",
        {"Inputs": {"A1": 4.0}, "Outputs": {"A1": "=Inputs!A1"}},
    )
    document = bindings_document(
        series_entry("value", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="key_scalar")
    assert pkg.compute_result.__key__ == ()
    assert pkg.compute_result.__domain__ == ((),)
    assert len(pkg.compute_result.__domain__) == len(pkg.compute_result(value=4.0))


def test_matrix_domain_is_row_major_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = _matrix_workbook(tmp_path)
    pkg = load_package(generate_inverted(workbook, _matrix_bindings()), tmp_path, name="key_matrix")
    assert pkg.data.REF_AREA_DOMAIN == ("France", "Kenya")
    assert pkg.data.TIME_PERIOD_DOMAIN == (2020, 2021)
    assert pkg.compute_ratio.__key__ == ("REF_AREA", "TIME_PERIOD")
    expected_domain = (
        ("France", 2020),
        ("France", 2021),
        ("Kenya", 2020),
        ("Kenya", 2021),
    )
    assert pkg.compute_ratio.__domain__ == expected_domain
    addresses = ["Engine!B8", "Engine!C8", "Engine!B9", "Engine!C9"]
    graph = create_dependency_graph(workbook, addresses, load_values=True)
    evaluated = FormulaEvaluator(graph).evaluate(addresses)
    got = pkg.compute_ratio()
    assert len(pkg.compute_ratio.__domain__) == len(got)
    for point, address in zip(pkg.compute_ratio.__domain__, addresses, strict=True):
        assert got[pkg.compute_ratio.__domain__.index(point)] == pytest.approx(evaluated[address])


def test_late_start_series_publishes_domain_slice(tmp_path: Path) -> None:
    workbook = _late_start_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _late_start_bindings()), tmp_path, name="key_late"
    )
    assert pkg.data.TIME_PERIOD_DOMAIN == (2020, 2021, 2022)
    assert pkg.compute_adj.__key__ == ("TIME_PERIOD",)
    assert pkg.compute_adj.__domain__ == (2021, 2022)
    assert pkg.compute_adj.__domain__ == pkg.data.TIME_PERIOD_DOMAIN[1:]
    assert pkg.internals.adj.__domain__ == (2021, 2022)
    result = pkg.compute_adj(values=pkg.data.VALUES_DEFAULT)
    assert len(pkg.compute_adj.__domain__) == len(result)
    assert result[pkg.compute_adj.__domain__.index(2021)] == pytest.approx(2.0)


def test_as_records_zips_key_and_domain(tmp_path: Path) -> None:
    years = (2008, 2009)
    workbook = _years_workbook(tmp_path, years)
    pkg = load_package(
        generate_inverted(workbook, _years_bindings(years)), tmp_path, name="key_records"
    )
    result = pkg.compute_path(growth=pkg.data.GROWTH_DEFAULT)
    records = pkg.as_records(pkg.compute_path, result)
    assert [row["TIME_PERIOD"] for row in records] == [2008, 2009]
    assert [row["OBS_VALUE"] for row in records] == pytest.approx((1.0, 2.0))
    matrix_pkg = load_package(
        generate_inverted(_matrix_workbook(tmp_path), _matrix_bindings()),
        tmp_path,
        name="key_records_matrix",
    )
    matrix_result = matrix_pkg.compute_ratio()
    matrix_records = matrix_pkg.as_records(matrix_pkg.compute_ratio, matrix_result)
    assert matrix_records[0]["REF_AREA"] == "France"
    assert matrix_records[0]["TIME_PERIOD"] == 2020
    assert matrix_records[-1]["REF_AREA"] == "Kenya"
    assert matrix_records[-1]["TIME_PERIOD"] == 2021


def test_domain_literals_stay_out_of_api_and_internals(tmp_path: Path) -> None:
    years = tuple(range(2008, 2014))
    modules = generate_inverted(_years_workbook(tmp_path, years), _years_bindings(years))
    assert "2008, 2009, 2010, 2011, 2012, 2013" in modules["data.py"]
    assert "2008, 2009, 2010, 2011, 2012, 2013" not in modules["api.py"]
    assert "2008, 2009, 2010, 2011, 2012, 2013" not in modules["internals.py"]
    assert "data.TIME_PERIOD_DOMAIN" in modules["api.py"]
    assert "data.TIME_PERIOD_DOMAIN" in modules["internals.py"]
