"""Layer A2 — internals take first-level deps, not leaves."""

from __future__ import annotations

from pathlib import Path
from typing import Literal

from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)


def _a2_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a2.xlsx",
        {
            "Inputs": {
                "A1": 1,
                "B1": -2,
                "C1": 2,
                "B10": 1,
                "C10": 2,
            },
            "Engine": {
                "A1": "=OFFSET(Inputs!B1,0,Inputs!A1-1)",
                "B1": "=3.5+CHOOSE(Inputs!A1,Engine!A1,0)",
                "C1": "=60*(1+0.04)/(1+Engine!B1/100)",
            },
            "Outputs": {
                "A1": "=Engine!C1",
            },
        },
    )


def _a2_bindings() -> dict:
    return bindings_document(
        series_entry("shock_type", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "shock_magnitudes",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
            key_concept="SHOCK_PARAMETER",
            key_read="int",
        ),
        series_entry(
            "shock_magnitude_resolved",
            "Engine!A1",
            layout="scalar",
            direction="internal",
        ),
        series_entry("shocked_growth", "Engine!B1", layout="scalar", direction="internal"),
        series_entry("path", "Engine!C1", layout="scalar", direction="internal"),
        series_entry("output_path", "Outputs!A1", layout="scalar", direction="output"),
    )


def _a2_dynamic_refs() -> DynamicRefConfig:
    return DynamicRefConfig.from_constraints({"Inputs!A1": Literal[1, 2]}, {})


def test_choose_and_offset_live_on_the_series_that_owns_them(tmp_path: Path) -> None:
    workbook = _a2_workbook(tmp_path)
    modules = generate_inverted(workbook, _a2_bindings(), dynamic_refs=_a2_dynamic_refs())
    internals = modules["internals.py"]
    assert "xl_choose" in internals
    assert "CHOOSE" in internals or "xl_choose" in internals
    # OFFSET is lowered to indexing on shock_magnitude_resolved, not on path.
    growth_fn_start = internals.index("def shocked_growth")
    path_fn_start = internals.index("def path")
    growth_body = internals[growth_fn_start:path_fn_start]
    path_body = internals[path_fn_start:]
    mag_body = internals[internals.index("def shock_magnitude_resolved") : growth_fn_start]
    assert "xl_choose" in growth_body
    assert "xl_choose" not in path_body
    assert "shock_magnitudes" in mag_body


def test_path_does_not_take_shock_type(tmp_path: Path) -> None:
    workbook = _a2_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _a2_bindings(), dynamic_refs=_a2_dynamic_refs()),
        tmp_path,
        name="a2_pkg",
    )
    growth_params = required_param_names(pkg.internals.shocked_growth)
    path_params = required_param_names(pkg.internals.path)
    mag_params = required_param_names(pkg.internals.shock_magnitude_resolved)
    assert "shock_type" in growth_params
    assert "shock_magnitude_resolved" in growth_params or any(
        "shock_magnitude" in name for name in growth_params
    )
    assert "path" not in growth_params
    assert "shock_type" not in path_params
    assert "shock_magnitudes" not in path_params
    assert "shocked_growth" in path_params
    assert "shock_type" in mag_params
    assert "shock_magnitudes" in mag_params
    assert "ctx" not in all_param_names(pkg.internals.path)
