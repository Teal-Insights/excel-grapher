"""Spike for #188: series bindings compile to the same `CellTypeEnv` as `constraints.py`."""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    EnumDomain,
    constraints_to_cell_type_env,
)
from excel_grapher.grapher.constraints import constraints_table, load_constraints_module
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefLimits
from excel_grapher.series_bindings import load_series_bindings
from excel_grapher.series_bindings.domains import (
    cell_type_env_from_bindings,
    dynamic_refs_from_bindings,
)
from tests.paths import INVERTED_TREE_TINY_DSA

# `input.domain` declarations equivalent to the fixture's constraints.py entries.
_TINY_DSA_INPUT_DOMAINS: dict[str, dict[str, Any]] = {
    "country_name": {"enum": ["Borvelia", "Litellia", "Aurelium"]},
    "shock_type": {"enum": [1, 2, 3]},
    "shock_year": {"between": {"min": 1, "max": 5}},
    "country_initial_debt": {"real_between": {"min": 0.0, "max": 200.0}},
    "shock_magnitudes": {"real_between": {"min": -30.0, "max": 30.0}},
    "growth_baseline": {"real_between": {"min": -10.0, "max": 15.0}},
    "interest_baseline": {"real_between": {"min": 0.0, "max": 20.0}},
    "primary_balance_baseline": {"real_between": {"min": -15.0, "max": 15.0}},
}


def _tiny_dsa_bindings_with_domains() -> Any:
    bindings = copy.deepcopy(load_series_bindings(INVERTED_TREE_TINY_DSA / "bindings"))
    for series in bindings["series"]:
        domain = _TINY_DSA_INPUT_DOMAINS.get(series["id"])
        if domain is not None:
            series["input"]["domain"] = domain
    return bindings


def test_tiny_dsa_bindings_compile_to_constraints_py_env() -> None:
    workbook = INVERTED_TREE_TINY_DSA / "tiny-dsa.xlsx"
    module = load_constraints_module(INVERTED_TREE_TINY_DSA / "constraints.py")
    expected = constraints_to_cell_type_env(constraints_table(module), {})

    env = cell_type_env_from_bindings(_tiny_dsa_bindings_with_domains(), workbook=workbook)

    assert env == expected


def test_dynamic_refs_from_bindings_wraps_env_and_limits() -> None:
    workbook = INVERTED_TREE_TINY_DSA / "tiny-dsa.xlsx"
    limits = DynamicRefLimits(max_branches=7)

    config = dynamic_refs_from_bindings(
        _tiny_dsa_bindings_with_domains(), workbook=workbook, limits=limits
    )

    assert isinstance(config, DynamicRefConfig)
    assert config.limits is limits
    assert config.cell_type_env["Inputs!B5"].enum == EnumDomain(
        values=frozenset({"Borvelia", "Litellia", "Aurelium"})
    )


def _scalar_series(series_id: str, address: str, **blocks: Any) -> dict[str, Any]:
    return {
        "id": series_id,
        "sheet": address.split("!")[0],
        "data_range": address,
        "layout": "scalar",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "string",
                "bind": {"kind": "data_cell", "read": "string"},
            },
            "dimensions": [],
        },
        "key": [],
        **blocks,
    }


def test_value_map_pins_cells_to_workbook_needles_and_unconstrained_inputs_are_skipped(
    tmp_path: Path,
) -> None:
    path = tmp_path / "book.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Dash")
    ws.write_string("B1", "GBR")
    ws.write_string("B2", "anything")
    ws.write_number("B3", 2.5)
    wb.close()
    bindings = {
        "schema_version": "1.16.0",
        "series": [
            _scalar_series(
                "country",
                "Dash!B1",
                input={"value_map": {"United Kingdom": "GBR", "France": "FRA"}},
            ),
            _scalar_series("free_text", "Dash!B2", input={}),
            _scalar_series("rate", "Dash!B3", constant={}),
        ],
    }

    env = cell_type_env_from_bindings(bindings, workbook=path)

    assert env == {
        "Dash!B1": CellType(
            kind=CellKind.STRING, enum=EnumDomain(values=frozenset({"GBR", "FRA"}))
        ),
        "Dash!B3": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2.5}))),
    }
