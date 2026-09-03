"""Layer A20 — formula matrices reading matrices join on the full key tuple (#612).

Since the key-point schedule join (#610), any `layout: matrix` series whose key
has a second dimension besides `TIME_PERIOD` failed closed on the identity
join: `_compute_preferred_fields` selected `("TIME_PERIOD",)`, so every
(country, year) cell shared its schedule coordinate with every other country at
the same year. The schedule coordinate is the full key tuple — `TIME_PERIOD`
as the inner schedule axis, the remaining key fields as the instance partition
— so a matrix is a loop nest and the identity join is unambiguous.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    write_workbook,
)


def _matrix_entry(
    series_id: str,
    data_range: str,
    *,
    header_row: int,
    label_column: str = "A",
    direction: str = "constant",
    key: list[str] | None = None,
) -> dict[str, Any]:
    sheet = data_range.split("!", 1)[0]
    dimensions: list[dict[str, Any]] = [
        {
            "id": "REF_AREA",
            "concept": "REF_AREA",
            "role": "key",
            "scope": "cell",
            "bind": {"kind": "row_label", "label_column": label_column, "read": "string"},
        },
        {
            "id": "TIME_PERIOD",
            "concept": "TIME_PERIOD",
            "role": "key",
            "scope": "cell",
            "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
        },
    ]
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
            "dimensions": dimensions,
        },
        "key": key if key is not None else ["REF_AREA", "TIME_PERIOD"],
    }
    if direction == "output":
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    elif direction == "internal":
        entry["internal"] = {}
    else:
        entry["constant"] = {}
    return entry


def _elementwise_workbook(tmp_path: Path) -> Path:
    """Country x year matrices; ratio = revenue / gdp, elementwise, no lag."""
    return write_workbook(
        tmp_path / "a20_elementwise.xlsx",
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


def _elementwise_bindings() -> dict:
    return bindings_document(
        _matrix_entry("gdp", "Engine!B2:C3", header_row=1),
        _matrix_entry("revenue", "Engine!B5:C6", header_row=4, key=["REF_AREA", "TIME_PERIOD"]),
        _matrix_entry("ratio", "Engine!B8:C9", header_row=7, direction="output"),
    )


def _zipper_workbook(tmp_path: Path) -> Path:
    """Matrix zipper: debt(c,t) = debt(c,t-1) + adj(c,t), adj = gdp * r."""
    return write_workbook(
        tmp_path / "a20_zipper.xlsx",
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
                "B5": "=B2*0.02",
                "C5": "=C2*0.02",
                "A6": "Kenya",
                "B6": "=B3*0.02",
                "C6": "=C3*0.02",
                "B7": 2020,
                "C7": 2021,
                "A8": "France",
                "B8": "=1000",
                "C8": "=B8+B5",
                "A9": "Kenya",
                "B9": "=500",
                "C9": "=B9+B6",
            },
        },
    )


def _zipper_bindings() -> dict:
    return bindings_document(
        _matrix_entry("gdp", "Engine!B2:C3", header_row=1),
        _matrix_entry("adjustment", "Engine!B5:C6", header_row=4, direction="internal"),
        _matrix_entry("debt", "Engine!B8:C9", header_row=7, direction="output"),
    )


def _evaluator_values(workbook: Path, addresses: list[str]) -> dict[str, object]:
    graph = create_dependency_graph(workbook, addresses, load_values=True)
    return FormulaEvaluator(graph).evaluate(addresses)


def test_elementwise_matrix_matches_evaluator(tmp_path: Path) -> None:
    workbook = _elementwise_workbook(tmp_path)
    modules = generate_inverted(workbook, _elementwise_bindings())
    pkg = load_package(modules, tmp_path, name="a20_elementwise")
    got = pkg.compute_ratio()
    addresses = ["Engine!B8", "Engine!C8", "Engine!B9", "Engine!C9"]
    expected = _evaluator_values(workbook, addresses)
    assert got == pytest.approx([expected[addr] for addr in addresses])


def test_matrix_zipper_matches_evaluator(tmp_path: Path) -> None:
    workbook = _zipper_workbook(tmp_path)
    modules = generate_inverted(workbook, _zipper_bindings())
    pkg = load_package(modules, tmp_path, name="a20_zipper")
    got = pkg.compute_debt()
    addresses = ["Engine!B8", "Engine!C8", "Engine!B9", "Engine!C9"]
    expected = _evaluator_values(workbook, addresses)
    assert got == pytest.approx([expected[addr] for addr in addresses])
