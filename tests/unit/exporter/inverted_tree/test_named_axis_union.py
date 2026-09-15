"""Issue 854 — subset axes share one master Axis instead of AXIS_2 clones.

Observed key tuples that are subsequences of a per-concept union share one
`Axis`. Series that do not cover the product use `Domain.explicit` /
`coordinate_runs` / `span`. Ragged Excel order that is not a subsequence still
clones.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.exporter.export_runtime.tensor import Axis
from excel_grapher.exporter.inverted_tree.named_axes import NamedAxes
from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    generate_inverted,
    write_workbook,
)


def test_named_axes_share_subsequence_keys_on_one_master() -> None:
    planned = NamedAxes.plan(
        (
            Axis("TIME_PERIOD", (2024, 2025, 2026), int),
            Axis("TIME_PERIOD", (2025, 2026), int),
            Axis("VARIANT", ("alpha",), str),
            Axis("VARIANT", ("beta",), str),
        )
    )
    names = [name for name, _axis in planned.items()]
    assert names == ["TIME_PERIOD_AXIS", "VARIANT_AXIS"]
    assert planned.constant(Axis("TIME_PERIOD", (2025, 2026), int)) == "TIME_PERIOD_AXIS"
    assert planned.emitted(Axis("TIME_PERIOD", (2025, 2026), int)).keys == (2024, 2025, 2026)
    assert planned.constant(Axis("VARIANT", ("beta",), str)) == "VARIANT_AXIS"
    assert planned.emitted(Axis("VARIANT", ("alpha",), str)).keys == ("alpha", "beta")


def test_named_axes_clone_ragged_order_that_is_not_a_subsequence() -> None:
    consecutive = Axis("TIME_PERIOD", (2024, 2025, 2026), int)
    ragged = Axis("TIME_PERIOD", (2025, 2026, 2024), int)
    planned = NamedAxes.plan((consecutive, ragged))
    assert [name for name, _axis in planned.items()] == ["TIME_PERIOD_AXIS", "TIME_PERIOD_AXIS_2"]
    assert planned.constant(consecutive) == "TIME_PERIOD_AXIS"
    assert planned.constant(ragged) == "TIME_PERIOD_AXIS_2"
    assert planned.emitted(consecutive).keys == (2024, 2025, 2026)
    assert planned.emitted(ragged).keys == (2025, 2026, 2024)


def _subset_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "axes.xlsx",
        {
            "Data": {
                "A1": "item",
                "B1": 2024,
                "C1": 2025,
                "D1": 2026,
                "A2": "alpha",
                "B2": 1,
                "C2": 2,
                "D2": 3,
                "A4": "item",
                "C4": 2025,
                "D4": 2026,
                "A5": "beta",
                "C5": 10,
                "D5": 11,
                "A7": "total",
                "B7": "=SUM(B2:D2)+SUM(C5:D5)",
            }
        },
    )


def _dim_row(column: str) -> dict[str, Any]:
    return {
        "id": "VARIANT",
        "concept": "VARIANT",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": column, "read": "string"},
    }


def _dim_col(header_row: int) -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _series(
    sid: str, data_range: str, key: list[str], dims: list[dict[str, Any]]
) -> dict[str, Any]:
    return {
        "id": sid,
        "sheet": "Data",
        "data_range": data_range,
        "layout": "series",
        "input": {},
        "key": key,
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": dims,
        },
    }


def _subset_bindings() -> dict[str, Any]:
    return {
        "schema_version": "1.16.0",
        "concept_scheme": {
            "id": "example",
            "concepts": [
                {"id": "OBS_VALUE", "dtype": "number"},
                {"id": "VARIANT", "dtype": "string"},
                {"id": "TIME_PERIOD", "dtype": "int"},
            ],
        },
        "series": [
            _series(
                "prices_full",
                "Data!B2:D2",
                ["VARIANT", "TIME_PERIOD"],
                [_dim_row("A"), _dim_col(1)],
            ),
            _series(
                "prices_short",
                "Data!C5:D5",
                ["VARIANT", "TIME_PERIOD"],
                [_dim_row("A"), _dim_col(4)],
            ),
            {
                "id": "total",
                "sheet": "Data",
                "data_range": "Data!B7",
                "layout": "scalar",
                "output": {"compute": {"name": "compute_total"}},
                "key": [],
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "float",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [],
                },
            },
        ],
    }


def test_subset_time_period_axes_share_one_constant(tmp_path: Path) -> None:
    modules = generate_inverted(_subset_workbook(tmp_path), _subset_bindings())
    data = modules["data.py"]
    internals = modules["internals.py"]
    assert data.count("TIME_PERIOD_AXIS = Axis(") == 1
    assert "TIME_PERIOD_AXIS_2" not in data
    assert data.count("VARIANT_AXIS = Axis(") == 1
    assert "VARIANT_AXIS_2" not in data
    assert "TIME_PERIOD_AXIS = Axis('TIME_PERIOD', (2024, 2025, 2026), int)" in data
    assert "VARIANT_AXIS = Axis('VARIANT', ('alpha', 'beta'), str)" in data
    assert (
        "PRICES_SHORT_DOMAIN = Domain.product("
        "Axis('VARIANT', ('beta',), str), "
        "Axis('TIME_PERIOD', span(TIME_PERIOD_AXIS, 2025, 2026), int))"
    ) in data
    assert "span(TIME_PERIOD_AXIS, 2025, 2026)" in data
    assert "VARIANT_AXIS_2" not in internals
    assert "TIME_PERIOD_AXIS_2" not in internals
    assert "span(data.TIME_PERIOD_AXIS, 2025, 2026)" in internals
    assert "rows=('beta',)" in internals
    pkg = assert_package_matches_evaluator(
        _subset_workbook(tmp_path), _subset_bindings(), tmp_path, "axis_union"
    )
    assert (
        pkg.compute_total(
            prices_full=pkg.data.PRICES_FULL_DEFAULT,
            prices_short=pkg.data.PRICES_SHORT_DEFAULT,
        )
        == 27
    )


def _ragged_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "ragged_axes.xlsx",
        {
            "Data": {
                "B1": 2024,
                "C1": 2025,
                "D1": 2026,
                "B2": 1,
                "C2": 2,
                "D2": 3,
                "B4": 2025,
                "C4": 2026,
                "D4": 2024,
                "B5": 10,
                "C5": 11,
                "D5": 12,
                "B7": "=SUM(B2:D2)+SUM(B5:D5)",
            }
        },
    )


def _time_series(sid: str, data_range: str, header_row: int) -> dict[str, Any]:
    return {
        "id": sid,
        "sheet": "Data",
        "data_range": data_range,
        "layout": "series",
        "input": {},
        "key": ["TIME_PERIOD"],
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [_dim_col(header_row)],
        },
    }


def _ragged_bindings() -> dict[str, Any]:
    return {
        "schema_version": "1.16.0",
        "concept_scheme": {
            "id": "example",
            "concepts": [
                {"id": "OBS_VALUE", "dtype": "number"},
                {"id": "TIME_PERIOD", "dtype": "int"},
            ],
        },
        "series": [
            _time_series("prices_full", "Data!B2:D2", 1),
            _time_series("prices_ragged", "Data!B5:D5", 4),
            {
                "id": "total",
                "sheet": "Data",
                "data_range": "Data!B7",
                "layout": "scalar",
                "output": {"compute": {"name": "compute_total"}},
                "key": [],
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "float",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [],
                },
            },
        ],
    }


def test_ragged_time_order_is_allowed_to_clone(tmp_path: Path) -> None:
    modules = generate_inverted(_ragged_workbook(tmp_path), _ragged_bindings())
    data = modules["data.py"]
    assert "TIME_PERIOD_AXIS = Axis('TIME_PERIOD', (2024, 2025, 2026), int)" in data
    assert "TIME_PERIOD_AXIS_2 = Axis('TIME_PERIOD', (2025, 2026, 2024), int)" in data
    pkg = assert_package_matches_evaluator(
        _ragged_workbook(tmp_path), _ragged_bindings(), tmp_path, "axis_ragged"
    )
    assert (
        pkg.compute_total(
            prices_full=pkg.data.PRICES_FULL_DEFAULT,
            prices_ragged=pkg.data.PRICES_RAGGED_DEFAULT,
        )
        == 39
    )
