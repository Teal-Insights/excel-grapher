"""Ranges over one bound series lower to lazy views, not per-cell callbacks."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)


def _lookup_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "lookup_view.xlsx",
        {
            "Lookup": {"A1": "AF", "A2": "BR", "A3": "KE", "B1": 1.0, "B2": 2.0, "B3": 3.0},
            "Inputs": {"A1": "BR"},
            "Outputs": {"A1": "=INDEX(Lookup!$B$1:$B$3, MATCH(Inputs!A1, Lookup!$A$1:$A$3, 0), 1)"},
        },
    )


def _lookup_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "codes",
            "Lookup!A1:A3",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "values",
            "Lookup!B1:B3",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry("code", "Inputs!A1", layout="scalar", direction="input", dtype="string"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_lookup_ranges_become_views_without_lambdas(tmp_path: Path) -> None:
    modules = generate_inverted(_lookup_workbook(tmp_path), _lookup_bindings())
    internals = modules["internals.py"]
    assert "lambda" not in internals
    assert "view(codes, rows=data.COUNTRY_AXIS.keys)" in internals
    assert "view(values, rows=data.COUNTRY_AXIS.keys)" in internals
    pkg = assert_package_matches_evaluator(
        _lookup_workbook(tmp_path), _lookup_bindings(), tmp_path, "lookup_view"
    )
    assert pkg.compute_out(values=pkg.data.VALUES_DEFAULT, code="KE") == 3.0


def _cumulative_workbook(tmp_path: Path, years: int) -> Path:
    cells: dict[str, object] = {}
    from fastpyxl.utils.cell import get_column_letter

    for index in range(years):
        column = get_column_letter(index + 2)
        cells[f"{column}1"] = 2020 + index
        cells[f"{column}2"] = float(index + 1)
        cells[f"{column}3"] = f"=SUM($B$2:{column}2)"
        cells[f"{column}4"] = f"=AVERAGE({get_column_letter(max(2, index))}2:{column}2)"
    return write_workbook(tmp_path / f"cumulative_{years}.xlsx", {"Sheet": cells})


def _cumulative_bindings(years: int) -> dict[str, Any]:
    from fastpyxl.utils.cell import get_column_letter

    last = get_column_letter(years + 1)
    return bindings_document(
        series_entry("flow", f"Sheet!B2:{last}2", layout="series", direction="input", header_row=1),
        series_entry(
            "cumulative", f"Sheet!B3:{last}3", layout="series", direction="output", header_row=1
        ),
        series_entry(
            "trailing", f"Sheet!B4:{last}4", layout="series", direction="output", header_row=1
        ),
    )


def test_growing_and_sliding_ranges_are_spans_over_the_axis(tmp_path: Path) -> None:
    modules = generate_inverted(_cumulative_workbook(tmp_path, 6), _cumulative_bindings(6))
    internals = modules["internals.py"]
    assert "span(data.TIME_PERIOD_AXIS, 2020, time_period)" in internals
    assert "span(data.TIME_PERIOD_AXIS, time_period - 2, time_period)" in internals
    assert_package_matches_evaluator(
        _cumulative_workbook(tmp_path, 6), _cumulative_bindings(6), tmp_path, "cumulative_6"
    )


def test_longer_horizon_does_not_replicate_formula_bodies(tmp_path: Path) -> None:
    sizes = {}
    for years in (10, 40):
        modules = generate_inverted(
            _cumulative_workbook(tmp_path, years), _cumulative_bindings(years)
        )
        sizes[years] = len(modules["internals.py"].encode())
    assert sizes[40] < sizes[10] * 1.25, sizes


def _keyed_scalar_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "keyed_scalar.xlsx",
        {
            "C1": {
                "A1": "AF",
                "B1": 1.0,
                "A2": "BR",
                "B2": 2.0,
                "A3": "KE",
                "B3": 3.0,
                "D1": "=SUM(B1:B3)",
            }
        },
    )


def _keyed_scalar_bindings() -> dict[str, Any]:
    total = {
        "id": "total",
        "sheet": "C1",
        "data_range": "C1!D1",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_total"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "SCENARIO",
                    "concept": "SCENARIO",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "sheet_name", "values": {"C1": "base"}},
                }
            ],
        },
        "key": ["SCENARIO"],
    }
    return bindings_document(
        series_entry(
            "cov",
            "C1!B1:B3",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        total,
    )


def test_keyed_scalar_hosts_have_no_loop_variables(tmp_path: Path) -> None:
    modules = generate_inverted(_keyed_scalar_workbook(tmp_path), _keyed_scalar_bindings())
    internals = modules["internals.py"]
    assert "_total_table_0 = view(cov, rows=data.COUNTRY_AXIS.keys)" in internals
    assert "scenario" not in internals
    pkg = assert_package_matches_evaluator(
        _keyed_scalar_workbook(tmp_path), _keyed_scalar_bindings(), tmp_path, "keyed_scalar"
    )
    assert pkg.compute_total(cov=pkg.data.COV_DEFAULT) == 6.0
