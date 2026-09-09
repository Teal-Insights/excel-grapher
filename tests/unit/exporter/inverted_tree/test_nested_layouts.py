"""Nested worksheet layouts keep grid provenance and product views.

A series whose rows enumerate two key fields (a group and a label inside
it) is neither a row nor a rectangle over one axis. Its provenance is a
grid of per-axis positions, and a range over its block is a product view
over the selected keys of each field.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.exporter.export_runtime.provenance import grid_cells
from excel_grapher.exporter.export_runtime.tensor import Axis, Domain, Tensor
from excel_grapher.exporter.inverted_tree import runtime
from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

COUNTRIES = Axis("COUNTRY", ("France", "Kenya"), str)
SCENARIOS = Axis("SCENARIO", ("base", "high", "low"), str)
YEARS = Axis("TIME_PERIOD", (2024, 2025, 2026), int)
NESTED = Domain.product(COUNTRIES, SCENARIOS, YEARS)


def test_grid_cells_compose_positions_per_axis_group() -> None:
    cells = grid_cells(
        "Vintage",
        NESTED,
        rows=(
            ("COUNTRY", "SCENARIO"),
            {
                ("France", "base"): 2,
                ("France", "high"): 3,
                ("France", "low"): 4,
                ("Kenya", "base"): 5,
                ("Kenya", "high"): 6,
                ("Kenya", "low"): 7,
            },
        ),
        cols=(("TIME_PERIOD",), {2024: "B", 2025: "C", 2026: "D"}),
    )
    assert cells[("Kenya", "high", 2025)] == "Vintage!C6"
    assert len(cells) == 18
    assert list(cells)[:2] == [("France", "base", 2024), ("France", "base", 2025)]
    assert ("Kenya", "high", 2027) not in cells
    fixed = grid_cells(
        "Sheet Two",
        Domain.product(YEARS),
        rows=9,
        cols=(("TIME_PERIOD",), {2024: "B", 2025: "C", 2026: "D"}),
        exceptions={(2026,): "Other!Z1"},
    )
    assert dict(fixed) == {
        (2024,): "'Sheet Two'!B9",
        (2025,): "'Sheet Two'!C9",
        (2026,): "Other!Z1",
    }


def test_product_view_enumerates_selected_keys_in_worksheet_order() -> None:
    values = Tensor(NESTED, tuple(float(i) for i in range(18)))
    block = runtime.view(
        values,
        rows={"COUNTRY": COUNTRIES.keys, "SCENARIO": runtime.span(SCENARIOS, "high", "low")},
        cols={"TIME_PERIOD": (2025,)},
    )
    assert block.shape == (4, 1)
    assert runtime.xl_sum(block) == 4.0 + 7.0 + 13.0 + 16.0
    assert runtime.xl_index(block, 3, 1) == 13.0


def _nested_workbook(tmp_path: Path) -> Path:
    cells: dict[str, object] = {"B1": 2024, "C1": 2025, "D1": 2026}
    row = 2
    for country in COUNTRIES.keys:
        for scenario in SCENARIOS.keys:
            cells[f"A{row}"] = scenario
            for offset, column in enumerate("BCD"):
                cells[f"{column}{row}"] = float(row * 10 + offset) + (
                    1.0 if country == "Kenya" else 0.0
                )
            row += 1
    for column in "BCD":
        cells[f"{column}9"] = f"=SUM({column}2:{column}7)"
        cells[f"{column}10"] = f"=SUM({column}5:{column}7)"
    return write_workbook(tmp_path / "nested.xlsx", {"Vintage": cells})


def _nested_bindings() -> dict[str, Any]:
    vintage = {
        "id": "vintage",
        "sheet": "Vintage",
        "data_range": "Vintage!B2:D7",
        "layout": "matrix",
        "input": {"setter": {"name": "set_vintage"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "COUNTRY",
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {
                        "kind": "value_map",
                        "values": {"France": "2:4", "Kenya": "5:7"},
                        "read": "string",
                    },
                },
                {
                    "id": "SCENARIO",
                    "concept": "SCENARIO",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["COUNTRY", "SCENARIO", "TIME_PERIOD"],
    }
    return bindings_document(
        vintage,
        series_entry("total", "Vintage!B9:D9", layout="series", direction="output", header_row=1),
        series_entry(
            "kenya_total", "Vintage!B10:D10", layout="series", direction="output", header_row=1
        ),
        schema_version="1.15.0",
    )


def test_nested_block_provenance_is_a_grid(tmp_path: Path) -> None:
    modules = generate_inverted(_nested_workbook(tmp_path), _nested_bindings())
    data = modules["data.py"]
    cells_line = next(line for line in data.splitlines() if line.startswith("VINTAGE_CELLS = "))
    assert cells_line.startswith("VINTAGE_CELLS = grid_cells('Vintage', VINTAGE_DOMAIN, ")
    assert "rows=(('COUNTRY', 'SCENARIO'), {('France', 'base'): 2," in cells_line
    assert "cols=(('TIME_PERIOD',), {2024: 'B', 2025: 'C', 2026: 'D'})" in cells_line
    assert "Vintage!C6" not in cells_line


def test_ranges_over_nested_blocks_are_product_views(tmp_path: Path) -> None:
    modules = generate_inverted(_nested_workbook(tmp_path), _nested_bindings())
    internals = modules["internals.py"]
    assert (
        "xl_sum(view(vintage, rows={'COUNTRY': data.COUNTRY_AXIS.keys, "
        "'SCENARIO': data.SCENARIO_AXIS.keys}, cols={'TIME_PERIOD': (time_period,)}))"
    ) in internals
    assert (
        "xl_sum(view(vintage, rows={'COUNTRY': ('Kenya',), "
        "'SCENARIO': data.SCENARIO_AXIS.keys}, cols={'TIME_PERIOD': (time_period,)}))"
    ) in internals
    assert "vintage['France', 'base', time_period]" not in internals
    pkg = assert_package_matches_evaluator(
        _nested_workbook(tmp_path), _nested_bindings(), tmp_path, "nested_views"
    )
    total = pkg.compute_total(vintage=pkg.data.VINTAGE_DEFAULT)
    assert total[2025] == sum(row * 10 + 1 for row in range(2, 8)) + 3.0
    assert dict(pkg.compute_total.__cells__) == {
        (2024,): "Vintage!B9",
        (2025,): "Vintage!C9",
        (2026,): "Vintage!D9",
    }
    assert pkg.data.VINTAGE_CELLS[("Kenya", "high", 2025)] == "Vintage!C6"
