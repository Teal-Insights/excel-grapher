"""Seed / terminal lag detection uses schedule coordinate + relative refs (#634).

Physical adjacency (`col ± 1` / `row + 1`) is the root cause of #631: an
absolute selector in the previous column is rewritten to `prior`. A seed or
terminal edge is a relative read of another bound series whose
`schedule_coord` is the host's coordinate ± 1.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import schedule_coord
from excel_grapher.exporter.inverted_tree.deps import (
    predecessor_address,
    successor_address,
)
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    oriented_addresses,
    oriented_document,
    series_entry,
    write_oriented_workbook,
)


def _selector_sheets() -> dict[str, dict[str, object]]:
    """#631 MCVE: absolute `$A$2` sits in the column before `B2:D2`."""
    return {
        "Engine": {
            "B1": 2009,
            "C1": 2010,
            "D1": 2011,
            "A2": "Nominal",
            "B2": '=IF($A$2=$A$5,B5,IF($A$2=$A$6,B6,""""))',
            "C2": '=IF($A$2=$A$5,C5,IF($A$2=$A$6,C6,""""))',
            "D2": '=IF($A$2=$A$5,D5,IF($A$2=$A$6,D6,""""))',
            "A5": "Nominal",
            "B5": 4.0,
            "C5": 5.0,
            "D5": 6.0,
            "A6": "Other",
            "B6": 1.0,
            "C6": 2.0,
            "D6": 3.0,
        },
    }


def _selector_bindings() -> dict:
    return bindings_document(
        series_entry("mode", "Engine!A2", layout="scalar", direction="input", dtype="string"),
        series_entry(
            "label_nominal",
            "Engine!A5",
            layout="scalar",
            direction="constant",
            dtype="string",
        ),
        series_entry(
            "label_other",
            "Engine!A6",
            layout="scalar",
            direction="constant",
            dtype="string",
        ),
        series_entry(
            "nominal",
            "Engine!B5:D5",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "other",
            "Engine!B6:D6",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "selected",
            "Engine!B2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _relative_seed_sheets() -> dict[str, dict[str, object]]:
    """Year-0 seed in the previous column; every debt formula is relative."""
    return {
        "Engine": {
            "A1": 2008,
            "B1": 2009,
            "C1": 2010,
            "D1": 2011,
            "A2": 100.0,
            "B2": "=A2*1.02",
            "C2": "=B2*1.02",
            "D2": "=C2*1.02",
        },
    }


def _relative_seed_bindings() -> dict:
    return bindings_document(
        series_entry(
            "seed",
            "Engine!A2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "debt",
            "Engine!B2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def _descending_seed_sheets() -> dict[str, dict[str, object]]:
    """Latest year leftmost; year-0 seed sits to the right of the series."""
    return {
        "Engine": {
            "A1": 2011,
            "B1": 2010,
            "C1": 2009,
            "D1": 2008,
            "A2": "=B2*1.02",
            "B2": "=C2*1.02",
            "C2": "=D2*1.02",
            "D2": 100.0,
        },
    }


def _descending_seed_bindings() -> dict:
    return bindings_document(
        series_entry(
            "seed",
            "Engine!D2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "debt",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def test_absolute_selector_is_not_a_scan_and_matches_evaluator(
    tmp_path: Path, orientation: str
) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a24_selector_{orientation[0]}.xlsx",
        _selector_sheets(),
        orientation=orientation,
    )
    document = oriented_document(_selector_bindings(), orientation)
    catalog, deps, _graph = inverted_graph_parts(workbook, document)
    assert deps["selected"].is_scan is False
    assert deps["selected"].seed_id is None
    assert predecessor_address(catalog.get("selected"), 0, catalog, _graph) is None

    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "prior" not in internals
    pkg = load_package(modules, tmp_path, name=f"a24_sel_{orientation[0]}")
    got = pkg.compute_selected(
        mode="Nominal",
        nominal=(4.0, 5.0, 6.0),
        other=(1.0, 2.0, 3.0),
    )
    cells = oriented_addresses(("Engine!B2", "Engine!C2", "Engine!D2"), orientation)
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, list(cells), load_values=True)
    ).evaluate(list(cells))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((4.0, 5.0, 6.0))


def test_relative_previous_period_seed_is_still_a_scan(tmp_path: Path, orientation: str) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a24_rel_{orientation[0]}.xlsx",
        _relative_seed_sheets(),
        orientation=orientation,
    )
    document = oriented_document(_relative_seed_bindings(), orientation)
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    debt = catalog.get("debt")
    assert deps["debt"].is_scan is True
    assert deps["debt"].seed_id == "seed"
    assert predecessor_address(debt, 0, catalog, graph) == catalog.get("seed").cells[0]

    pkg = load_package(
        generate_inverted(workbook, document), tmp_path, name=f"a24_rel_{orientation[0]}"
    )
    cells = oriented_addresses(("Engine!B2", "Engine!C2", "Engine!D2"), orientation)
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, list(cells), load_values=True)
    ).evaluate(list(cells))
    assert pkg.compute_debt(seed=100.0) == pytest.approx(tuple(expected[cell] for cell in cells))
    assert pkg.compute_debt(seed=100.0) == pytest.approx((102.0, 104.04, 106.1208))


def test_descending_seed_classifies_via_schedule_coord(tmp_path: Path, orientation: str) -> None:
    workbook = write_oriented_workbook(
        tmp_path / f"a24_desc_{orientation[0]}.xlsx",
        _descending_seed_sheets(),
        orientation=orientation,
    )
    document = oriented_document(_descending_seed_bindings(), orientation)
    catalog, deps, _graph = inverted_graph_parts(workbook, document)
    debt = catalog.get("debt")
    assert deps["debt"].seed_id == "seed"
    assert deps["debt"].is_scan is True
    year0 = min(range(len(debt.cells)), key=lambda i: schedule_coord(debt.cells[i], catalog))
    assert year0 != 0

    pkg = load_package(
        generate_inverted(workbook, document), tmp_path, name=f"a24_desc_{orientation[0]}"
    )
    cells = oriented_addresses(("Engine!A2", "Engine!B2", "Engine!C2"), orientation)
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, list(cells), load_values=True)
    ).evaluate(list(cells))
    got = pkg.compute_debt(seed=100.0)
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((106.1208, 104.04, 102.0))


def test_successor_terminal_uses_schedule_not_orientation_guess(tmp_path: Path) -> None:
    """Vertical terminal below the series: no `is_vertical` row arithmetic."""
    workbook = write_oriented_workbook(
        tmp_path / "a24_term.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "C1": 2011,
                "D1": 2012,
                "A2": "=B2*0.9",
                "B2": "=C2*0.9",
                "C2": "=D2*0.9",
                "D2": 100.0,
            },
        },
        orientation="vertical",
    )
    document = oriented_document(
        bindings_document(
            series_entry(
                "terminal",
                "Engine!D2",
                layout="series",
                direction="input",
                header_row=1,
            ),
            series_entry(
                "value",
                "Engine!A2:C2",
                layout="series",
                direction="output",
                header_row=1,
            ),
        ),
        "vertical",
    )
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    value = catalog.get("value")
    assert deps["value"].scan_direction == "reversed"
    assert deps["value"].seed_id == "terminal"
    last = max(range(len(value.cells)), key=lambda i: schedule_coord(value.cells[i], catalog))
    assert successor_address(value, last, catalog, graph) == catalog.get("terminal").cells[0]
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a24_term")
    assert pkg.compute_value(terminal=100.0) == pytest.approx((72.9, 81.0, 90.0))
