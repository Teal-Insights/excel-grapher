"""Issue 799 — compact mostly regular sparse multidimensional domains."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    make_catalog,
    series_entry,
    write_workbook,
)


def _country_dim() -> dict[str, Any]:
    return {
        "id": "COUNTRY",
        "concept": "COUNTRY",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
    }


def _time_dim(*, header_row: int = 1) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _matrix_entry(
    series_id: str,
    data_range: str | list[str],
    *,
    direction: str,
    header_row: int = 1,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "Data",
        "data_range": data_range,
        "layout": "matrix",
        "key": ["COUNTRY", "TIME_PERIOD"],
        "structure": {
            "measure": _measure(),
            "dimensions": [_country_dim(), _time_dim(header_row=header_row)],
        },
    }
    if direction == "input":
        entry["input"] = {}
    else:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    return entry


def _result_entry(size: int, *, last_row: int | None = None) -> dict[str, Any]:
    end = last_row if last_row is not None else size + 1
    return {
        "id": "result",
        "sheet": "Data",
        "data_range": f"Data!E2:E{end}",
        "layout": "series",
        "output": {"compute": {"name": "compute_result"}},
        "key": ["COUNTRY"],
        "structure": {
            "measure": _measure(),
            "dimensions": [_country_dim()],
        },
    }


def _missing_corner_workbook(tmp_path: Path, size: int) -> Path:
    cells: dict[str, object] = {"B1": 2020, "C1": 2021}
    last = size + 1
    for row in range(2, last + 1):
        cells[f"A{row}"] = f"Country {row}"
        cells[f"B{row}"] = float(row)
        cells[f"C{row}"] = float(row * 2)
        cells[f"E{row}"] = f"=B{row}+C{row}"
    cells[f"E{last}"] = f"=B{last}"
    return write_workbook(tmp_path / f"missing_corner_{size}.xlsx", {"Data": cells})


def _missing_corner_bindings(size: int) -> dict[str, Any]:
    last = size + 1
    return bindings_document(
        _matrix_entry("values", [f"Data!B2:C{last - 1}", f"Data!B{last}"], direction="input"),
        _result_entry(size),
    )


def _bound_matrix(
    series_id: str,
    points: tuple[tuple[object, ...], ...],
    *,
    keys: tuple[str, ...] = ("COUNTRY", "TIME_PERIOD"),
) -> BoundSeries:
    cells = tuple(f"Data!A{index + 1}" for index in range(len(points)))
    domain = tuple(KeyPoint(tuple(zip(keys, point, strict=True))) for point in points)
    return BoundSeries(
        series_id=series_id,
        layout="matrix",
        direction="input",
        cells=cells,
        key_fields=keys,
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement(series_id, series_id, None, 0, len(cells), cells, domain),),
    )


def _catalog_for(*series: BoundSeries):
    mapping = {item.series_id: item for item in series}
    order = tuple(item.series_id for item in series)
    address_to_id = {cell: item.series_id for item in series for cell in item.cells}
    return make_catalog(mapping, order, address_to_id)


def test_sparse_domain_source_and_import_scale(tmp_path: Path) -> None:
    sizes = (10, 40)
    data_sizes: list[int] = []
    domain_sizes: list[int] = []
    for size in sizes:
        workbook = _missing_corner_workbook(tmp_path, size)
        document = _missing_corner_bindings(size)
        modules = generate_inverted(workbook, document)
        data_py = modules["data.py"]
        interned_line = next(
            line
            for line in data_py.splitlines()
            if "Domain.product(COUNTRY_AXIS)" in line or line.startswith("RESULT_DOMAIN =")
        )
        data_sizes.append(len(data_py))
        domain_sizes.append(len(interned_line))
        assert "Country " not in interned_line
    assert domain_sizes[-1] / domain_sizes[0] < 2.0
    assert data_sizes[-1] / data_sizes[0] < 3.0


def _ragged_labelled_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "ragged_labelled.xlsx",
        {
            "Origin": {"A1": 2017},
            "Data": {
                "B1": "=Origin!A1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "A4": "outstanding",
                "B4": 10.0,
                "C4": 20.0,
                "D4": 30.0,
                "A5": "short",
                "B5": 1.0,
                "C5": 2.0,
                "E4": "=B4+C4+D4+B5+C5",
            },
        },
    )


def _ragged_labelled_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("first_year", "Origin!A1", direction="input", dtype="int"),
        series_entry(
            "year_labels",
            "Data!B1:D1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        {
            "id": "stock",
            "sheet": "Data",
            "data_range": ["Data!B4:D4", "Data!B5:C5"],
            "layout": "matrix",
            "constant": {},
            "key": ["COUNTRY", "TIME_PERIOD"],
            "structure": {
                "measure": _measure(),
                "dimensions": [_country_dim(), _time_dim()],
            },
        },
        series_entry("total", "Data!E4", direction="output"),
        schema_version="1.17.0",
    )


def test_runtime_labelled_sparse_constant_imports(tmp_path: Path) -> None:
    """A labelled constant keeps the cells it binds, including a missing corner.

    The snapshot domain passed to `define_series` must match those cells.
    A product of the axis keys counts the missing corner and the import fails.
    """
    modules = generate_inverted(_ragged_labelled_workbook(tmp_path), _ragged_labelled_bindings())
    pkg = load_package(modules, tmp_path, "ragged_const")
    stock = pkg.data.STOCK
    assert len(stock.domain) == 5
    assert ("short", 2019) not in set(stock.domain)
    assert stock[("outstanding", 2017)] == 10.0
    assert stock[("outstanding", 2019)] == 30.0
    assert stock[("short", 2018)] == 2.0
    assert stock[("short", 2019)] is None
