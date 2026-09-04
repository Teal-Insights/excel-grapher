"""Layer A20 — formula matrices reading matrices join on the full key tuple (#612).

Since the key-point schedule join (#610), any `layout: matrix` series whose key
has a second dimension besides `TIME_PERIOD` failed closed on the identity
join: `_compute_preferred_fields` selected `("TIME_PERIOD",)`, so every
(country, year) cell shared its schedule coordinate with every other country at
the same year. The schedule coordinate is the full key tuple — `TIME_PERIOD`
as the inner schedule axis, the remaining key fields as the instance partition
— so a matrix is a loop nest and the identity join is unambiguous.

#638: flattened tuple order makes a late-start `adj` block non-contiguous
(`[1, 2, 4, 5]`), so fusion must check contiguity per outer-key block and
emit an outer loop over the instance partition.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import schedule_coord
from excel_grapher.exporter.inverted_tree.deps import collect_all_dependence_edges
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import plan_fused_scc, plan_scc
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    oriented_addresses,
    oriented_document,
    write_oriented_workbook,
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


def _zipper_sheets() -> dict[str, dict[str, object]]:
    """Country × year zipper matching #638: adj starts one year late.

    Flattened `(REF_AREA, TIME_PERIOD)` coords are `debt [0..5]` and
    `adj [1, 2, 4, 5]` — non-contiguous across countries, contiguous
    within each outer-key block.
    """
    return {
        "Engine": {
            "B1": 2020,
            "C1": 2021,
            "D1": 2022,
            "A2": "France",
            "B2": "=1000",
            "C2": "=B2+C5",
            "D2": "=C2+D5",
            "A3": "Kenya",
            "B3": "=500",
            "C3": "=B3+C6",
            "D3": "=C3+D6",
            "B4": 2020,
            "C4": 2021,
            "D4": 2022,
            "A5": "France",
            "C5": "=B2*0.02",
            "D5": "=C2*0.02",
            "A6": "Kenya",
            "C6": "=B3*0.02",
            "D6": "=C3*0.02",
        }
    }


def _zipper_workbook(tmp_path: Path) -> Path:
    """Matrix zipper: debt(c,t) = debt(c,t-1) + adj(c,t), adj = debt(c,t-1) * r."""
    return write_workbook(tmp_path / "a20_zipper.xlsx", _zipper_sheets())


def _zipper_bindings() -> dict:
    return bindings_document(
        _matrix_entry("adjustment", "Engine!C5:D6", header_row=4, direction="internal"),
        _matrix_entry("debt", "Engine!B2:D3", header_row=1, direction="output"),
    )


def _zipper_debt_addresses() -> list[str]:
    return ["Engine!B2", "Engine!C2", "Engine!D2", "Engine!B3", "Engine!C3", "Engine!D3"]


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


@pytest.mark.parametrize("orientation", ["horizontal", "vertical"])
def test_matrix_zipper_emits_rung_2_and_matches_evaluator(tmp_path: Path, orientation: str) -> None:
    sheets = _zipper_sheets()
    workbook = write_oriented_workbook(
        tmp_path / f"a20_zipper_{orientation}.xlsx", sheets, orientation=orientation
    )
    document = oriented_document(_zipper_bindings(), orientation)
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    choice = plan_scc(("debt", "adjustment"), catalog=catalog, graph=graph)
    assert choice.rung == 2
    assert choice.plan is not None
    adj_coords = [schedule_coord(cell, catalog) for cell in catalog.get("adjustment").cells]
    assert sorted(adj_coords) == [1, 2, 4, 5]
    assert adj_coords != list(range(min(adj_coords), max(adj_coords) + 1))
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "scan_debt_adjustment" in internals or "scan_adjustment_debt" in internals
    pkg = load_package(modules, tmp_path, name=f"a20_zip_{orientation[:1]}")
    addresses = list(oriented_addresses(_zipper_debt_addresses(), orientation))
    expected = _evaluator_values(workbook, addresses)
    assert pkg.compute_debt() == pytest.approx(tuple(expected[addr] for addr in addresses))


def _sized_zipper_sheets(n_areas: int, n_years: int) -> dict[str, dict[str, object]]:
    cells: dict[str, object] = {}
    for year_i in range(n_years):
        cells[f"{get_column_letter(year_i + 2)}1"] = 2020 + year_i
    for area_i in range(n_areas):
        row = 2 + area_i
        cells[f"A{row}"] = f"C{area_i:02d}"
        cells[f"B{row}"] = f"={1000 + 100 * area_i}"
        adj_row = 3 + n_areas + area_i
        for year_i in range(1, n_years):
            col = get_column_letter(year_i + 2)
            pred = get_column_letter(year_i + 1)
            cells[f"{col}{row}"] = f"={pred}{row}+{col}{adj_row}"
    adj_header = 2 + n_areas
    for year_i in range(n_years):
        cells[f"{get_column_letter(year_i + 2)}{adj_header}"] = 2020 + year_i
    for area_i in range(n_areas):
        row = 3 + n_areas + area_i
        cells[f"A{row}"] = f"C{area_i:02d}"
        debt_row = 2 + area_i
        for year_i in range(1, n_years):
            col = get_column_letter(year_i + 2)
            pred = get_column_letter(year_i + 1)
            cells[f"{col}{row}"] = f"={pred}{debt_row}*0.02"
    return {"Engine": cells}


def _sized_zipper_bindings(n_areas: int, n_years: int) -> dict[str, Any]:
    last_col = get_column_letter(n_years + 1)
    last_debt_row = 1 + n_areas
    adj_header = 2 + n_areas
    last_adj_row = adj_header + n_areas
    return bindings_document(
        _matrix_entry(
            "adjustment",
            f"Engine!C{adj_header + 1}:{last_col}{last_adj_row}",
            header_row=adj_header,
            direction="internal",
        ),
        _matrix_entry(
            "debt",
            f"Engine!B2:{last_col}{last_debt_row}",
            header_row=1,
            direction="output",
        ),
    )


def test_matrix_zipper_code_size_independent_of_countries_and_years(tmp_path: Path) -> None:
    small_wb = write_workbook(
        tmp_path / "a20_size_small.xlsx", _sized_zipper_sheets(n_areas=2, n_years=3)
    )
    large_wb = write_workbook(
        tmp_path / "a20_size_large.xlsx", _sized_zipper_sheets(n_areas=5, n_years=8)
    )
    small = generate_inverted(small_wb, _sized_zipper_bindings(2, 3))
    large = generate_inverted(large_wb, _sized_zipper_bindings(5, 8))
    for filename in ("api.py", "internals.py"):
        small_lines = small[filename].splitlines()
        large_lines = large[filename].splitlines()
        assert len(small_lines) == len(large_lines), (
            f"{filename} grew from {len(small_lines)} to {len(large_lines)} lines"
        )
        assert abs(len(small[filename]) - len(large[filename])) <= 48


def _cross_country_legal_sheets() -> dict[str, dict[str, object]]:
    """Kenya@t reads France@t; France is the first catalog partition."""
    return {
        "Engine": {
            "B1": 2020,
            "C1": 2021,
            "A2": "France",
            "B2": "=100",
            "C2": "=B2+1",
            "A3": "Kenya",
            "B3": "=B2",
            "C3": "=C2",
        }
    }


def _cross_country_legal_bindings() -> dict[str, Any]:
    return bindings_document(
        _matrix_entry("path", "Engine!B2:C3", header_row=1, direction="output"),
    )


def test_legal_cross_country_read_fuses(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a20_cross_ok.xlsx", _cross_country_legal_sheets())
    document = _cross_country_legal_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    edges = collect_all_dependence_edges(catalog, graph)
    cross = [
        edge
        for edge in edges
        if edge.consumer_cell == "Engine!B3" and edge.producer_cell == "Engine!B2"
    ]
    assert cross and all(edge.access == "cross_partition" for edge in cross)
    plan = plan_fused_scc(("path",), catalog=catalog, graph=graph)
    assert plan is not None
    assert plan_scc(("path",), catalog=catalog, graph=graph).rung != 3
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a20_cross_ok")
    addresses = ["Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3"]
    expected = _evaluator_values(workbook, addresses)
    assert pkg.compute_path() == pytest.approx(tuple(expected[addr] for addr in addresses))


def _cross_country_cycle_sheets() -> dict[str, dict[str, object]]:
    return {
        "Engine": {
            "B1": 2020,
            "C1": 2021,
            "A2": "France",
            "B2": "=B3",
            "C2": "=C3",
            "A3": "Kenya",
            "B3": "=B2",
            "C3": "=C2",
        }
    }


def test_mutual_cross_country_same_index_fails_closed_naming_cells(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a20_cross_cycle.xlsx", _cross_country_cycle_sheets())
    document = _cross_country_legal_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    with pytest.raises(InvertedTreeExportError, match="Engine!B2") as exc:
        plan_fused_scc(("path",), catalog=catalog, graph=graph)
    message = str(exc.value)
    assert "Engine!B3" in message
    with pytest.raises(InvertedTreeExportError, match="Engine!B2"):
        generate_inverted(workbook, document)
