"""Issue 949 — catalog Domain from `row_index` / `column_letter` binds."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.deps import _key_field_axis
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from tests.unit.exporter.inverted_tree.helpers import (
    transpose_bindings,
    write_workbook,
)


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _row_value_map(values: dict[int, Any]) -> dict[str, Any]:
    return {
        "id": "ROW",
        "concept": "ROW",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "value_map", "values": values, "read": "int"},
    }


def _col_value_map(values: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": "COL",
        "concept": "COL",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "value_map", "values": values, "read": "string"},
    }


def _row_index_dim() -> dict[str, Any]:
    return {
        "id": "ROW",
        "concept": "ROW",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_index", "read": "int"},
    }


def _column_letter_dim() -> dict[str, Any]:
    return {
        "id": "COL",
        "concept": "COL",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_letter", "read": "string"},
    }


def _pin_series(*, values_row: dict[int, Any] | None, values_col: dict[str, Any] | None) -> dict:
    if values_row is None:
        row_dim = _row_index_dim()
    else:
        row_dim = _row_value_map(values_row)
    if values_col is None:
        col_dim = _column_letter_dim()
    else:
        col_dim = _col_value_map(values_col)
    return {
        "id": "pins",
        "sheet": "Engine",
        "data_range": ["Engine!B2:C2", "Engine!B4"],
        "layout": "series",
        "constant": {},
        "structure": {"measure": _measure(), "dimensions": [row_dim, col_dim]},
        "key": ["ROW", "COL"],
    }


def _bindings(series: dict[str, Any]) -> dict[str, Any]:
    return validate_bindings_document(
        {"schema_version": CURRENT_SCHEMA_VERSION, "series": [series]}
    )


def _grab_bag_workbook(tmp_path: Path) -> Path:
    return write_workbook(tmp_path / "mcve.xlsx", {"Engine": {"B2": 1, "C2": 2, "B4": 3}})


def test_address_binds_match_identity_value_maps(tmp_path: Path) -> None:
    path = _grab_bag_workbook(tmp_path)
    identity = build_catalog(
        _bindings(_pin_series(values_row={2: 2, 4: 4}, values_col={"B": "B", "C": "C"})),
        workbook=path,
    ).get("pins")
    compact = build_catalog(
        _bindings(_pin_series(values_row=None, values_col=None)),
        workbook=path,
    ).get("pins")

    assert tuple(compact.tensor_domain) == tuple(identity.tensor_domain)
    assert dict(compact.coordinate_cells) == dict(identity.coordinate_cells)
    assert set(compact.coordinate_cells) == {(2, "B"), (2, "C"), (4, "B")}
    assert compact.coordinate_cells[(2, "B")].endswith("!B2")
    assert compact.coordinate_cells[(2, "C")].endswith("!C2")
    assert compact.coordinate_cells[(4, "B")].endswith("!B4")


def test_row_index_alone_still_collides_on_shared_row(tmp_path: Path) -> None:
    path = _grab_bag_workbook(tmp_path)
    series = {
        "id": "collapsed_rows",
        "sheet": "Engine",
        "data_range": ["Engine!B2:C2", "Engine!B4"],
        "layout": "series",
        "constant": {},
        "structure": {
            "measure": _measure(),
            "dimensions": [_row_index_dim()],
        },
        "key": ["ROW"],
    }
    catalog = build_catalog(_bindings(series), workbook=path)
    with pytest.raises(
        InvertedTreeExportError,
        match="domain contains duplicate coordinates; correct the authored keys",
    ):
        _ = catalog.get("collapsed_rows").tensor_domain


def test_address_bind_kinds_declare_row_and_column_axes(tmp_path: Path) -> None:
    path = _grab_bag_workbook(tmp_path)
    series = build_catalog(
        _bindings(_pin_series(values_row=None, values_col=None)),
        workbook=path,
    ).get("pins")
    assert _key_field_axis(series, "ROW") == "row"
    assert _key_field_axis(series, "COL") == "col"


def test_transpose_swaps_row_index_and_column_letter() -> None:
    document = {
        "schema_version": CURRENT_SCHEMA_VERSION,
        "series": [
            {
                "id": "pins",
                "data_range": "Engine!B2:C2",
                "structure": {
                    "dimensions": [
                        {"bind": {"kind": "row_index", "read": "int"}},
                        {"bind": {"kind": "column_letter", "read": "string"}},
                    ]
                },
            }
        ],
    }
    transposed = transpose_bindings(document)
    binds = [dim["bind"] for dim in transposed["series"][0]["structure"]["dimensions"]]
    assert binds[0]["kind"] == "column_letter"
    assert binds[1]["kind"] == "row_index"
    restored = transpose_bindings(transposed)
    restored_binds = [dim["bind"] for dim in restored["series"][0]["structure"]["dimensions"]]
    assert restored_binds[0]["kind"] == "row_index"
    assert restored_binds[1]["kind"] == "column_letter"
