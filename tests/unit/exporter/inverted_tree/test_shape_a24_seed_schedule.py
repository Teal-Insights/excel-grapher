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


def test_two_unkeyed_scalars_are_not_an_ambiguous_seed(tmp_path: Path) -> None:
    """A formula that reads two scalars has no unique seed (a5 shocked_path)."""
    workbook = write_oriented_workbook(
        tmp_path / "a24_two_scalars.xlsx",
        {
            "Inputs": {"A1": 10.0, "B1": 2},
            "Engine": {
                "C4": 1,
                "D4": 2,
                "C5": 1,
                "D5": 2,
                "C7": "=Inputs!A1+IF(Engine!C5>=Inputs!B1,1,0)",
                "D7": "=Inputs!A1+IF(Engine!D5>=Inputs!B1,1,0)",
            },
        },
        orientation="horizontal",
    )
    document = bindings_document(
        series_entry("value", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("shock_year", "Inputs!B1", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "engine_year_labels",
            "Engine!C5:D5",
            layout="series",
            direction="constant",
            dtype="int",
            header_row=4,
        ),
        series_entry(
            "shocked_path",
            "Engine!C7:D7",
            layout="series",
            direction="output",
            header_row=5,
        ),
    )
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert predecessor_address(catalog.get("shocked_path"), 0, catalog, graph) is None
    assert deps["shocked_path"].seed_id is None
    assert deps["shocked_path"].is_scan is False


def _nest_entry(series_id: str, data_range: str, *, header_row: int, label_column: str) -> dict:
    sheet = data_range.split("!", 1)[0]
    return {
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
            "dimensions": [
                {
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": label_column, "read": "string"},
                },
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
                },
            ],
        },
        "key": ["COUNTRY", "TIME_PERIOD"],
        "input": {"setter": {"name": f"set_{series_id}"}},
    }


def test_two_schedule_adjacent_seeds_are_not_a_unique_seed(tmp_path: Path) -> None:
    """Two producers at host − 1 degrade to no seed, not a first-wins pick."""
    workbook = write_oriented_workbook(
        tmp_path / "a24_two_seeds.xlsx",
        {
            "Engine": {
                "B1": 2008,
                "C1": 2009,
                "A2": "US",
                "A3": "EU",
                "B2": 10.0,
                "B3": 20.0,
                "C2": "=B2+B3",
            },
        },
        orientation="horizontal",
    )
    document = bindings_document(
        _nest_entry("seed_a", "Engine!B2", header_row=1, label_column="A"),
        _nest_entry("seed_b", "Engine!B3", header_row=1, label_column="A"),
        series_entry("path", "Engine!C2", layout="series", direction="output", header_row=1),
    )
    with pytest.warns(UserWarning, match="ambiguous seed"):
        _catalog, deps, _graph = inverted_graph_parts(workbook, document)
        modules = generate_inverted(workbook, document)
    assert deps["path"].seed_id is None
    pkg = load_package(modules, tmp_path, name="a24_two_seeds")
    got = pkg.compute_path(seed_a=10.0, seed_b=20.0)
    assert got == pytest.approx((30.0,))
