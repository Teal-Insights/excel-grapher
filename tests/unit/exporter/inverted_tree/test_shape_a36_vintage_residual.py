"""Layer A36 — vintage residual orientation can flip across ISSUANCE_YEAR (#762).

Two `ISSUANCE_YEAR × TIME_PERIOD` series are a DAG inside each vintage, but
the distance-zero residual used to union those edges at one `TIME_PERIOD`.
Vintage 2024 amortizes (`stock → principal`); vintage 2025 opens from this
year's issuance (`principal → stock`). That is a per-partition DAG, not an
Excel circular and not a must-cycle.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.deps import collect_all_dependence_edges
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    plan_scc,
    residual_body_order,
)
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    generate_inverted,
    inverted_graph_parts,
    load_package,
    write_workbook,
)

_MEASURE = {
    "concept": "OBS_VALUE",
    "dtype": "float",
    "bind": {"kind": "data_cell", "read": "float"},
}


def _time() -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
    }


def _issuance() -> dict[str, Any]:
    return {
        "id": "ISSUANCE_YEAR",
        "concept": "ISSUANCE_YEAR",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": "A", "read": "int"},
    }


def vintage_residual_workbook(tmp_path: Path, stem: str = "a36_vintage") -> Path:
    """Two vintages, two years: 2024 amortizes in 2025; 2025 opens from issuance."""
    return write_workbook(
        tmp_path / f"{stem}.xlsx",
        {
            "V": {
                "D1": 2024,
                "E1": 2025,
                "A5": 2024,
                "A6": 2025,
                "A7": 2024,
                "A8": 2025,
                "D5": 100,
                "E5": "=D5-E7",
                "D6": 0,
                "E6": 50,
                "D7": 0,
                "E7": "=IF(E$1>2024,$D5/2,0)",
                "D8": 0,
                "E8": "=E6",
                "A10": "=E5",
            }
        },
    )


def _keyed(series_id: str, data_range: str, *, output: bool = False) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "V",
        "data_range": data_range,
        "layout": "series",
        "structure": {
            "measure": _MEASURE,
            "dimensions": [_issuance(), _time()],
        },
        "key": ["ISSUANCE_YEAR", "TIME_PERIOD"],
    }
    if output:
        entry["output"] = {"compute": {"name": f"compute_{series_id}"}}
    else:
        entry["internal"] = {}
    return entry


def vintage_residual_bindings() -> dict[str, Any]:
    return {
        "schema_version": "1.14.0",
        "concept_scheme": {
            "id": "a36_vintage",
            "concepts": [
                {"id": "TIME_PERIOD", "dtype": "int"},
                {"id": "OBS_VALUE", "dtype": "number"},
                {"id": "ISSUANCE_YEAR", "dtype": "int"},
            ],
        },
        "series": [
            {
                "id": "years",
                "sheet": "V",
                "data_range": "V!D1:E1",
                "layout": "series",
                "constant": {},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "int",
                        "bind": {"kind": "data_cell", "read": "int"},
                    },
                    "dimensions": [_time()],
                },
                "key": ["TIME_PERIOD"],
            },
            _keyed("stock", "V!D5:E6", output=True),
            _keyed("principal", "V!D7:E8"),
            {
                "id": "result",
                "sheet": "V",
                "data_range": "V!A10",
                "layout": "scalar",
                "output": {"compute": {"name": "compute_result"}},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "float",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [],
                },
                "key": [],
            },
        ],
    }


def _same_vintage_cycle_workbook(tmp_path: Path) -> Path:
    """Vintage 2024 has a real same-year stock ⇄ principal cycle."""
    return write_workbook(
        tmp_path / "a36_same_vintage_cycle.xlsx",
        {
            "V": {
                "D1": 2024,
                "E1": 2025,
                "A5": 2024,
                "A6": 2025,
                "A7": 2024,
                "A8": 2025,
                "D5": 100,
                "E5": "=E7",
                "D6": 0,
                "E6": 50,
                "D7": 0,
                "E7": "=E5",
                "D8": 0,
                "E8": "=E6",
                "A10": "=E5",
            }
        },
    )


def _evaluator_values(workbook: Path, addresses: list[str]) -> dict[str, object]:
    graph = create_dependency_graph(workbook, addresses, load_values=True)
    return FormulaEvaluator(graph).evaluate(addresses)


def test_vintage_residual_is_legal_per_issuance_year(tmp_path: Path) -> None:
    workbook = vintage_residual_workbook(tmp_path)
    document = vintage_residual_bindings()
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    choice = plan_scc(("stock", "principal"), catalog=catalog, graph=graph)
    assert choice.rung == 2
    plan = choice.plan
    assert plan is not None
    assert plan.is_nested
    assert plan.unroll
    assert plan.partitions == ((2024,), (2025,))
    assert plan.partition_regions[0][-1].body_order == ("principal", "stock")
    assert plan.partition_regions[1][-1].body_order == ("stock", "principal")
    edges = collect_all_dependence_edges(catalog, graph)
    assert residual_body_order(("stock", "principal"), edges, catalog) is None


def test_vintage_residual_export_matches_evaluator(tmp_path: Path) -> None:
    workbook = vintage_residual_workbook(tmp_path)
    document = vintage_residual_bindings()
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert "eval_instance" not in internals
    pkg = load_package(modules, tmp_path, name="a36_vintage")
    stock_cells = ["V!D5", "V!E5", "V!D6", "V!E6"]
    expected = _evaluator_values(workbook, [*stock_cells, "V!A10"])
    assert dict(pkg.compute_stock().items()) == pytest.approx(
        {coordinate: expected[cell] for coordinate, cell in pkg.data.STOCK_CELLS.items()}
    )
    result = pkg.compute_result()
    if isinstance(result, tuple):
        assert len(result) == 1
        result = result[0]
    assert result == pytest.approx(expected["V!A10"])


def test_same_vintage_distance_zero_cycle_still_fail_closed(tmp_path: Path) -> None:
    workbook = _same_vintage_cycle_workbook(tmp_path)
    document = vintage_residual_bindings()
    with pytest.raises(InvertedTreeExportError, match="distance-zero residual"):
        generate_inverted(workbook, document)
