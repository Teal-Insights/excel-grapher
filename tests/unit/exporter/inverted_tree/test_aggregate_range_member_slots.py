"""Literal aggregate ranges resolve catalog slots per statement member (#775).

Sparse catalogs have irregular flattened slots, so classifying a `SUM`
column window through the affine block-axis origin fails closed. Dense tables
must not freeze the first member's `take` indices for every host cell. In-SCC
producers demand each selected instance instead of reading an incomplete tuple.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)

_SPARSE_TOTALS = (11.0, 122.0, 1033.0)
_DENSE_TOTALS = (16.0, 122.0, 1033.0)
_SCC_TOTALS = (66.0, 55.0, 33.0)


def _time_dim() -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
    }


def _country_dim() -> dict[str, Any]:
    return {
        "id": "COUNTRY",
        "concept": "COUNTRY",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
    }


def _grid_entry(
    series_id: str,
    data_range: str | list[str],
    *,
    direction: str = "constant",
) -> dict[str, Any]:
    if isinstance(data_range, list):
        sheet = data_range[0].split("!", 1)[0]
    else:
        sheet = data_range.split("!", 1)[0]
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": sheet,
        "data_range": data_range,
        "layout": "series",
        "key": ["COUNTRY", "TIME_PERIOD"],
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [_time_dim(), _country_dim()],
        },
    }
    if direction == "constant":
        entry["constant"] = {}
    elif direction == "internal":
        entry["internal"] = {}
    else:
        raise ValueError(f"unknown direction {direction!r}")
    return entry


def _totals_entry() -> dict[str, Any]:
    return series_entry(
        "totals",
        "Engine!B6:D6",
        layout="series",
        direction="output",
        header_row=1,
        compute_name="compute_totals",
    )


def _column_sum_cells(*, b3: object, b4: object, c4: object, d4: object) -> dict[str, object]:
    return {
        "A2": "a",
        "A3": "b",
        "A4": "c",
        "B1": 2020,
        "C1": 2021,
        "D1": 2022,
        "B2": 1,
        "C2": 2,
        "D2": 3,
        "B3": b3,
        "C3": 20,
        "D3": 30,
        "B4": b4,
        "C4": c4,
        "D4": d4,
        "B6": "=SUM(B2:B4)",
        "C6": "=SUM(C2:C4)",
        "D6": "=SUM(D2:D4)",
    }


def _sparse_workbook(tmp_path: Path) -> Path:
    cells = _column_sum_cells(b3=None, b4=10, c4=100, d4=1000)
    del cells["B3"]
    return write_workbook(tmp_path / "agg_sparse.xlsx", {"Engine": cells})


def _dense_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "agg_dense.xlsx",
        {"Engine": _column_sum_cells(b3=5, b4=10, c4=100, d4=1000)},
    )


def _scc_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "agg_scc.xlsx",
        {"Engine": _column_sum_cells(b3=10, b4="=C6", c4="=D6", d4=0)},
    )


def _sparse_bindings() -> dict[str, Any]:
    return bindings_document(
        _grid_entry("vintages", ["Engine!B2:D2", "Engine!C3:D3", "Engine!B4:D4"]),
        _totals_entry(),
    )


def _dense_bindings() -> dict[str, Any]:
    return bindings_document(_grid_entry("vintages", "Engine!B2:D4"), _totals_entry())


def _scc_bindings() -> dict[str, Any]:
    return bindings_document(
        _grid_entry("vintages", "Engine!B2:D4", direction="internal"),
        _totals_entry(),
    )


@pytest.mark.parametrize("force_rung", [None, 3])
def test_sparse_column_sum_exports_and_matches_column_totals(
    tmp_path: Path, force_rung: Literal[3] | None
) -> None:
    workbook = _sparse_workbook(tmp_path)
    modules = generate_inverted(
        workbook,
        _sparse_bindings(),
        force_rung=force_rung,
        blank_ranges=["Engine!B3"],
    )
    pkg = load_package(modules, tmp_path, name=f"agg_sparse_{force_rung}")
    assert pkg.compute_totals() == pytest.approx(_SPARSE_TOTALS)


@pytest.mark.parametrize("force_rung", [None, 3])
def test_dense_column_sum_does_not_freeze_first_column(
    tmp_path: Path, force_rung: Literal[3] | None
) -> None:
    workbook = _dense_workbook(tmp_path)
    modules = generate_inverted(workbook, _dense_bindings(), force_rung=force_rung)
    internals = modules["internals.py"]
    pkg = load_package(modules, tmp_path, name=f"agg_dense_{force_rung}")
    got = pkg.compute_totals()
    assert got != pytest.approx((_DENSE_TOTALS[0],) * 3)
    assert got == pytest.approx(_DENSE_TOTALS)
    assert "take(" in internals


def test_sparse_column_sum_does_not_use_affine_origin_classifier(tmp_path: Path) -> None:
    generate_inverted(
        _sparse_workbook(tmp_path),
        _sparse_bindings(),
        blank_ranges=["Engine!B3"],
    )


@pytest.mark.parametrize("force_rung", [None, 3])
def test_in_scc_column_sum_demands_selected_instances(
    tmp_path: Path, force_rung: Literal[3] | None
) -> None:
    workbook = _scc_workbook(tmp_path)
    modules = generate_inverted(workbook, _scc_bindings(), force_rung=force_rung)
    pkg = load_package(modules, tmp_path, name=f"agg_scc_{force_rung}")
    assert pkg.compute_totals() == pytest.approx(_SCC_TOTALS)
    if force_rung == 3:
        assert "demand_instance(" in modules["internals.py"]
