"""IF dual-reads of two SCENARIO pairs in one producer stay keyed (#752).

`IF($selector, path_a[t], path_b[t])` is a same-year dual on `SCENARIO`.
When one producer holds two such pairs (B2.1/B2.2 and B6.1/B6.2), each host
row dual-reads its own pair. `_host_follow_key_maps` cannot remap `SCENARIO`
(two leftovers per member), literal pattern sets disagree, and export
fail-closes.

Correspondence is IF then/else operand position, not catalog order or string
prefixes. Emit indexes a host-ordered pair table; `TIME_PERIOD` follows the
host. A dual without an IF stays unclassifiable.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
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

_PATH_SCENARIO = {
    "id": "SCENARIO",
    "concept": "SCENARIO",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "value_map",
        "values": {
            "B2.1 Market": 3,
            "B2.2 Non-Market": 4,
            "B6.1 Market": 5,
            "B6.2 Non-Market": 6,
        },
        "read": "string",
    },
}

_HOST_SCENARIO = {
    "id": "SCENARIO",
    "concept": "SCENARIO",
    "role": "key",
    "scope": "cell",
    "bind": {
        "kind": "value_map",
        "values": {"B2": 10, "B6": 11},
        "read": "string",
    },
}


def _measure(*, dtype: str = "float") -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": dtype,
        "bind": {"kind": "data_cell", "read": dtype},
    }


def _unwrap(value: object) -> object:
    """Return the sole member of a 1-tuple compute result."""
    if isinstance(value, tuple) and len(value) == 1:
        return value[0]
    return value


def _mcve_sheets(
    *,
    c10: str = "=IF($G$1=1,C3,C4)",
    d10: str = "=IF($G$1=1,D3,D4)",
    c11: str = "=IF($G$1=1,C5,C6)",
    d11: str = "=IF($G$1=1,D5,D6)",
    flag: object = 1,
) -> dict[str, dict[str, object]]:
    return {
        "Engine": {
            "C1": 2024,
            "D1": 2025,
            "C3": 10,
            "D3": 11,
            "C4": 20,
            "D4": 21,
            "C5": 30,
            "D5": 31,
            "C6": 40,
            "D6": 41,
            "G1": flag,
            "C10": c10,
            "D10": d10,
            "C11": c11,
            "D11": d11,
        }
    }


def _paths_entry() -> dict[str, Any]:
    return {
        "id": "paths",
        "sheet": "Engine",
        "data_range": "Engine!C3:D6",
        "layout": "matrix",
        "input": {"setter": {"name": "set_paths"}},
        "structure": {
            "measure": _measure(),
            "dimensions": [_PATH_SCENARIO, _TIME_DIM],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }


def _selected_entry() -> dict[str, Any]:
    return {
        "id": "selected",
        "sheet": "Engine",
        "data_range": ["Engine!C10:D10", "Engine!C11:D11"],
        "layout": "series",
        "output": {"compute": {"name": "compute_selected"}},
        "structure": {
            "measure": _measure(),
            "dimensions": [_HOST_SCENARIO, _TIME_DIM],
        },
        "key": ["SCENARIO", "TIME_PERIOD"],
    }


def _flag_entry(*, dtype: str = "float") -> dict[str, Any]:
    return {
        "id": "flag",
        "sheet": "Engine",
        "data_range": "Engine!G1",
        "layout": "scalar",
        "input": {"setter": {"name": "set_flag"}},
        "structure": {"measure": _measure(dtype=dtype), "dimensions": []},
        "key": [],
    }


def _mcve_bindings(*, flag_dtype: str = "float") -> dict[str, Any]:
    return bindings_document(
        _paths_entry(),
        _selected_entry(),
        _flag_entry(dtype=flag_dtype),
        schema_version="1.14.0",
    )


def test_scenario_pair_dual_read_is_keyed(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a34.xlsx", _mcve_sheets())
    catalog, deps, _graph = inverted_graph_parts(workbook, _mcve_bindings())
    host = deps["selected"]
    assert "paths" in host.param_ids
    assert "paths" in host.keyed_ids
    assert "paths" not in host.lagged_ids
    assert "paths" not in host.aligned_ids
    assert catalog.get("paths").cells == (
        "Engine!C3",
        "Engine!D3",
        "Engine!C4",
        "Engine!D4",
        "Engine!C5",
        "Engine!D5",
        "Engine!C6",
        "Engine!D6",
    )


def test_scenario_pair_dual_read_emits_and_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "a34_eval.xlsx", _mcve_sheets())
    document = _mcve_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "paths" in deps["selected"].keyed_ids
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    assert (
        "(('B2.1 Market', 'B2.2 Non-Market'), ('B2.1 Market', 'B2.2 Non-Market'), "
        "('B6.1 Market', 'B6.2 Non-Market'), ('B6.1 Market', 'B6.2 Non-Market'))[i][0]"
    ) in internals
    assert (
        "(('B2.1 Market', 'B2.2 Non-Market'), ('B2.1 Market', 'B2.2 Non-Market'), "
        "('B6.1 Market', 'B6.2 Non-Market'), ('B6.1 Market', 'B6.2 Non-Market'))[i][1]"
    ) in internals
    pkg = load_package(modules, tmp_path, name="a34_eval")
    cells = ["Engine!C10", "Engine!D10", "Engine!C11", "Engine!D11"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = call_compute(pkg, "selected", kwargs)
    assert got == pytest.approx((10.0, 11.0, 30.0, 31.0))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    flipped = call_compute(pkg, "selected", {**kwargs, "flag": 0})
    assert flipped == pytest.approx((20.0, 21.0, 40.0, 41.0))


def test_scenario_pair_dual_read_follows_if_branch_order(tmp_path: Path) -> None:
    """THEN/ELSE alignment wins over catalog row order (#752)."""
    sheets = _mcve_sheets(
        c10="=IF($G$1=1,C4,C3)",
        d10="=IF($G$1=1,D4,D3)",
        c11="=IF($G$1=1,C5,C6)",
        d11="=IF($G$1=1,D5,D6)",
    )
    workbook = write_workbook(tmp_path / "a34_swap.xlsx", sheets)
    document = _mcve_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "paths" in deps["selected"].keyed_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a34_swap")
    kwargs = input_kwargs(catalog, graph)
    assert call_compute(pkg, "selected", kwargs) == pytest.approx((20.0, 21.0, 30.0, 31.0))
    assert call_compute(pkg, "selected", {**kwargs, "flag": 0}) == pytest.approx(
        (10.0, 11.0, 40.0, 41.0)
    )


def test_scenario_pair_dual_read_per_row_if_condition(tmp_path: Path) -> None:
    """Q-CRAFT-like `IF($selector=label, then, else)` still aligns on branches."""
    sheets = _mcve_sheets(
        c10='=IF($G$1="B2.1 Market",C3,C4)',
        d10='=IF($G$1="B2.1 Market",D3,D4)',
        c11='=IF($G$1="B6.1 Market",C5,C6)',
        d11='=IF($G$1="B6.1 Market",D5,D6)',
        flag="B2.1 Market",
    )
    workbook = write_workbook(tmp_path / "a34_qcraft.xlsx", sheets)
    document = _mcve_bindings(flag_dtype="string")
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "paths" in deps["selected"].keyed_ids
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="a34_qcraft")
    kwargs = input_kwargs(catalog, graph)
    assert _unwrap(kwargs["flag"]) == "B2.1 Market"
    assert call_compute(pkg, "selected", kwargs) == pytest.approx((10.0, 11.0, 40.0, 41.0))
    assert call_compute(pkg, "selected", {**kwargs, "flag": "B6.1 Market"}) == pytest.approx(
        (20.0, 21.0, 30.0, 31.0)
    )


def test_scenario_pair_sum_without_if_is_unclassifiable(tmp_path: Path) -> None:
    sheets = _mcve_sheets(
        c10="=C3+C4",
        d10="=D3+D4",
        c11="=C5+C6",
        d11="=D5+D6",
    )
    workbook = write_workbook(tmp_path / "a34_sum.xlsx", sheets)
    with pytest.raises(InvertedTreeExportError, match=r"Engine!C10.*B2\.1 Market") as exc:
        inverted_graph_parts(workbook, _mcve_bindings())
    message = str(exc.value)
    assert "Engine!C3" in message
    assert "Engine!C4" in message
    assert "B2.2 Non-Market" in message
