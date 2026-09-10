"""Formula families fold across coordinates instead of unrolling per cell.

A reference that points at the same cell from every host cell is a literal
key even when it was authored without `$`. An integer key of another axis
that moves with the host period is the period variable plus a difference.
A label key that embeds the host's own key is a template over it.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from fastpyxl.utils.cell import get_column_letter

from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)


def _cumulative_workbook(tmp_path: Path, years: int) -> Path:
    cells: dict[str, object] = {}
    for index in range(years):
        column = get_column_letter(index + 2)
        cells[f"{column}1"] = 2020 + index
        cells[f"{column}2"] = float(index + 1)
        cells[f"{column}3"] = f"=SUM(B2:{column}2)"
    return write_workbook(tmp_path / "cumulative_relative.xlsx", {"Sheet": cells})


def _cumulative_bindings(years: int) -> dict[str, Any]:
    last = get_column_letter(years + 1)
    return bindings_document(
        series_entry("flow", f"Sheet!B2:{last}2", layout="series", direction="input", header_row=1),
        series_entry(
            "cumulative", f"Sheet!B3:{last}3", layout="series", direction="output", header_row=1
        ),
    )


def test_relative_range_start_pinned_to_the_first_column_is_literal(tmp_path: Path) -> None:
    modules = generate_inverted(_cumulative_workbook(tmp_path, 6), _cumulative_bindings(6))
    internals = modules["internals.py"]
    assert "span(data.TIME_PERIOD_AXIS, 2020, time_period)" in internals
    assert "time_period ==" not in internals
    assert_package_matches_evaluator(
        _cumulative_workbook(tmp_path, 6), _cumulative_bindings(6), tmp_path, "cumulative_rel"
    )


def _vintage_workbook(tmp_path: Path) -> Path:
    cells: dict[str, object] = {"B1": 2024, "C1": 2025, "D1": 2026}
    for row, issued in enumerate((2024, 2025, 2026), start=2):
        cells[f"A{row}"] = issued
        for offset, column in enumerate("BCD"):
            cells[f"{column}{row}"] = float(row * 10 + offset)
    for row, column in enumerate("BCD", start=2):
        cells[f"{column}6"] = f"=SUM({column}$2:{column}{row})"
    return write_workbook(tmp_path / "vintage.xlsx", {"Vintage": cells})


def _vintage_bindings() -> dict[str, Any]:
    vintage = {
        "id": "vintage",
        "sheet": "Vintage",
        "data_range": "Vintage!B2:D4",
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
                    "id": "ISSUANCE_YEAR",
                    "concept": "ISSUANCE_YEAR",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "int"},
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
        "key": ["ISSUANCE_YEAR", "TIME_PERIOD"],
    }
    document = bindings_document(
        vintage,
        series_entry(
            "outstanding", "Vintage!B6:D6", layout="series", direction="output", header_row=1
        ),
        schema_version="1.15.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "ISSUANCE_YEAR", "dtype": "int"})
    return document


def test_triangular_range_end_follows_the_host_period(tmp_path: Path) -> None:
    modules = generate_inverted(_vintage_workbook(tmp_path), _vintage_bindings())
    internals = modules["internals.py"]
    assert (
        "xl_sum(view(vintage, rows=span(data.ISSUANCE_YEAR_AXIS, 2024, time_period), "
        "cols=(time_period,)))"
    ) in internals
    assert "time_period ==" not in internals
    pkg = assert_package_matches_evaluator(
        _vintage_workbook(tmp_path), _vintage_bindings(), tmp_path, "vintage_triangle"
    )
    outstanding = pkg.compute_outstanding(vintage=pkg.data.VINTAGE_DEFAULT)
    assert outstanding[2026] == 22.0 + 32.0 + 42.0


def _terms_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "terms.xlsx",
        {
            "Terms": {"A1": "Grace France", "B1": 5.0, "A2": "Grace Kenya", "B2": 3.0},
            "Model": {"A1": "France", "B1": "=Terms!B1*2", "A2": "Kenya", "B2": "=Terms!B2*2"},
        },
    )


def _terms_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "terms",
            "Terms!B1:B2",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="SCENARIO",
            key_read="string",
        ),
        series_entry(
            "doubled",
            "Model!B1:B2",
            layout="series",
            direction="output",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
    )


def test_label_keys_built_from_the_host_key_are_templates(tmp_path: Path) -> None:
    modules = generate_inverted(_terms_workbook(tmp_path), _terms_bindings())
    internals = modules["internals.py"]
    assert "terms[f'Grace {country}']" in internals
    assert not re.search(r"if country ==", internals)
    pkg = assert_package_matches_evaluator(
        _terms_workbook(tmp_path), _terms_bindings(), tmp_path, "terms_template"
    )
    assert pkg.compute_doubled(terms=pkg.data.TERMS_DEFAULT)["Kenya"] == 6.0
