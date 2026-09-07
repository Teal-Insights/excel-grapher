"""T+1 IF look-ahead into a richer-keyed producer is not an ambiguous seed (#745).

A host keyed by `TIME_PERIOD` that reads `producer[holder, t+1]` through
`IF(cond, A+B, C)` is one take of a holder nest, not competing scan
terminals. `unique_seed_or_none` must not fail-close on several keyed
cells at `host ± 1`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.deps import successor_address
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    write_workbook,
)

_TIME_DIM = {
    "id": "TIME_PERIOD",
    "concept": "TIME_PERIOD",
    "role": "key",
    "scope": "cell",
    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
}


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _holder_dim(values: dict[str, int]) -> dict[str, Any]:
    return {
        "id": "HOLDER",
        "concept": "HOLDER",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "value_map", "values": values, "read": "string"},
    }


def _mcve_sheets() -> dict[str, dict[str, object]]:
    return {
        "Engine": {
            "B1": 2010,
            "C1": 2011,
            "D1": 2012,
            "A2": "residents",
            "B2": 10,
            "C2": 11,
            "D2": 12,
            "A3": "non-residents",
            "B3": 20,
            "C3": 21,
            "D3": 22,
            "A4": "non-residents",
            "B4": 30,
            "C4": 31,
            "D4": 32,
            "G1": 1,
            "B5": "=IF($G$1=1,C3+C4,C2)",
            "C5": "=IF($G$1=1,D3+D4,D2)",
        },
    }


def _mcve_bindings() -> dict[str, Any]:
    document = bindings_document(
        {
            "id": "fx_st",
            "sheet": "Engine",
            "data_range": "Engine!B2:D4",
            "layout": "series",
            "exclude_rows": ["3"],
            "input": {"setter": {"name": "set_fx_st"}},
            "structure": {
                "measure": _measure(),
                "dimensions": [
                    _holder_dim({"residents": 2, "non-residents": 4}),
                    _TIME_DIM,
                ],
            },
            "key": ["HOLDER", "TIME_PERIOD"],
        },
        {
            "id": "lc_st",
            "sheet": "Engine",
            "data_range": "Engine!B3:D3",
            "layout": "series",
            "input": {"setter": {"name": "set_lc_st"}},
            "structure": {
                "measure": _measure(),
                "dimensions": [_holder_dim({"non-residents": 3}), _TIME_DIM],
            },
            "key": ["HOLDER", "TIME_PERIOD"],
        },
        {
            "id": "flag",
            "sheet": "Engine",
            "data_range": "Engine!G1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_flag"}},
            "structure": {"measure": _measure(), "dimensions": []},
            "key": [],
        },
        {
            "id": "stock",
            "sheet": "Engine",
            "data_range": "Engine!B5:C5",
            "layout": "series",
            "output": {"compute": {"name": "compute_stock"}},
            "structure": {"measure": _measure(), "dimensions": [_TIME_DIM]},
            "key": ["TIME_PERIOD"],
        },
        schema_version="1.14.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "HOLDER", "dtype": "string"})
    return document


def test_if_lookahead_into_richer_key_is_not_an_ambiguous_seed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a34.xlsx", _mcve_sheets())
    with pytest.warns(UserWarning, match="ambiguous seed"):
        catalog, deps, graph = inverted_graph_parts(workbook, _mcve_bindings())
    host = catalog.get("stock")
    assert deps["stock"].seed_id is None
    assert successor_address(host, 0, catalog, graph) == host.cells[1]


def test_if_lookahead_into_richer_key_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a34_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    with pytest.warns(UserWarning, match="ambiguous seed"):
        catalog, deps, graph = inverted_graph_parts(workbook, document)
        modules = generate_inverted(workbook, document)
    assert deps["stock"].seed_id is None
    pkg = load_package(modules, tmp_path, name="a34_eval")
    cells = ["Engine!B5", "Engine!C5"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = call_compute(pkg, "stock", input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((52.0, 54.0))
