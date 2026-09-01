"""Pass-1 shape-unit suite (issue #595) — RED until series-helper collapse lands.

Each MCVE pins one observable contract for binding-named, key-parameterized
helpers. Current excel-grapher still emits per-cell ``cell_*`` functions, so
these assertions fail for the right reasons until
``CodeGenerator.generate_modules`` collapses bound series.
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document
from tests.integration.exporter.pass1_shape_contract import (
    assert_compute_calls_helper,
    assert_helper_inventory,
    assert_helper_signature,
    assert_no_cell_defs_for_addresses,
    def_names,
)


def _time_period_dim(*, header_row: int = 1) -> dict[str, Any]:
    return {
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": "int"},
    }


def _helper_body(internals: str, name: str) -> str:
    from tests.integration.exporter.pass1_shape_contract import function_source

    return function_source(internals, name)


def _generate(
    workbook: Path,
    *,
    bindings_doc: dict[str, Any],
    targets: list[str],
) -> dict[str, str]:
    bindings = validate_bindings_document(deepcopy(bindings_doc))
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def _write_row_series_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    engine = wb.add_worksheet("Engine")
    for col, year in enumerate([1, 2, 3, 4, 5], start=3):
        engine.write(0, col - 1, year)
        engine.write_number(4, col - 1, float(year))
        engine.write_formula(9, col - 1, f"=IF(Engine!{chr(64 + col)}5>=Inputs!$B$1,1,0)")
    inputs = wb.add_worksheet("Inputs")
    inputs.write_number("B1", 3)
    wb.close()


_ROW_SERIES_BINDINGS: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "row_series.xlsx",
    "series": [
        {
            "id": "shock_year",
            "sheet": "Inputs",
            "data_range": "Inputs!B1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_shock_year"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "int",
                    "bind": {"kind": "data_cell", "read": "int"},
                },
                "dimensions": [],
            },
            "key": [],
        },
        {
            "id": "shock_flag",
            "sheet": "Engine",
            "data_range": "Engine!C10:G10",
            "layout": "series",
            "internal": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "int",
                    "bind": {"kind": "data_cell", "read": "int"},
                },
                "dimensions": [_time_period_dim(header_row=1)],
            },
            "key": ["TIME_PERIOD"],
            "validation": {"intersect_graph_formulas": True},
        },
    ],
}


def test_row_series_time_period_emits_one_helper(tmp_path: Path) -> None:
    workbook = tmp_path / "row_series.xlsx"
    _write_row_series_workbook(workbook)
    targets = expand_data_range("Engine!C10:G10", workbook=workbook) + ["Inputs!B1"]
    files = _generate(workbook, bindings_doc=_ROW_SERIES_BINDINGS, targets=targets)
    internals = files["internals.py"]

    assert_helper_inventory(internals, {"shock_flag"})
    assert_helper_signature(internals, "shock_flag", ("ctx", "time_period"))
    assert_no_cell_defs_for_addresses(
        internals,
        [f"Engine!{col}10" for col in "CDEFG"],
    )
    assert "time_period" in _helper_body(internals, "shock_flag")


def _write_scalar_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    inputs = wb.add_worksheet("Inputs")
    inputs.write_number("B1", 40.0)
    engine = wb.add_worksheet("Engine")
    engine.write_formula("B6", "=Inputs!B1")
    wb.close()


_SCALAR_BINDINGS: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "scalar.xlsx",
    "series": [
        {
            "id": "seed_debt",
            "sheet": "Inputs",
            "data_range": "Inputs!B1",
            "layout": "scalar",
            "constant": {},
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
        {
            "id": "resolved_debt",
            "sheet": "Engine",
            "data_range": "Engine!B6",
            "layout": "scalar",
            "internal": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [],
            },
            "key": [],
            "validation": {"intersect_graph_formulas": True},
        },
    ],
}


def test_scalar_internal_helper(tmp_path: Path) -> None:
    workbook = tmp_path / "scalar.xlsx"
    _write_scalar_workbook(workbook)
    files = _generate(
        workbook,
        bindings_doc=_SCALAR_BINDINGS,
        targets=["Engine!B6", "Inputs!B1"],
    )
    internals = files["internals.py"]
    assert_helper_inventory(internals, {"resolved_debt"})
    assert_helper_signature(internals, "resolved_debt", ("ctx",))
    assert_no_cell_defs_for_addresses(internals, ["Engine!B6"])


def _write_recurrence_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    engine = wb.add_worksheet("Engine")
    for col, year in enumerate([1, 2, 3], start=3):
        engine.write(0, col - 1, year)
    inputs = wb.add_worksheet("Inputs")
    inputs.write_number("B1", 50.0)
    inputs.write_number("B2", 0.05)
    engine.write_formula("C6", "=Inputs!B1*(1+Inputs!B2)")
    engine.write_formula("D6", "=C6*(1+Inputs!B2)")
    engine.write_formula("E6", "=D6*(1+Inputs!B2)")
    wb.close()


_RECURRENCE_BINDINGS: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "recurrence.xlsx",
    "series": [
        {
            "id": "initial_debt",
            "sheet": "Inputs",
            "data_range": "Inputs!B1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_initial_debt"}},
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
        {
            "id": "growth",
            "sheet": "Inputs",
            "data_range": "Inputs!B2",
            "layout": "scalar",
            "input": {"setter": {"name": "set_growth"}},
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
        {
            "id": "path",
            "sheet": "Engine",
            "data_range": "Engine!C6:E6",
            "layout": "series",
            "internal": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [_time_period_dim(header_row=1)],
            },
            "key": ["TIME_PERIOD"],
            "validation": {"intersect_graph_formulas": True},
        },
    ],
}


def test_self_recurrence_year_anchor(tmp_path: Path) -> None:
    workbook = tmp_path / "recurrence.xlsx"
    _write_recurrence_workbook(workbook)
    targets = expand_data_range("Engine!C6:E6", workbook=workbook) + [
        "Inputs!B1",
        "Inputs!B2",
    ]
    files = _generate(workbook, bindings_doc=_RECURRENCE_BINDINGS, targets=targets)
    internals = files["internals.py"]

    assert_helper_inventory(internals, {"path"})
    assert_helper_signature(internals, "path", ("ctx", "time_period"))
    assert_no_cell_defs_for_addresses(
        internals,
        ["Engine!C6", "Engine!D6", "Engine!E6"],
    )
    body = _helper_body(internals, "path")
    assert "time_period" in body
    assert "time_period == 1" in body or "time_period==1" in body
    assert "path(ctx, time_period=time_period - 1)" in body or (
        "path(ctx, time_period=time_period-1)" in body
    )


def _write_compute_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    engine = wb.add_worksheet("Engine")
    outputs = wb.add_worksheet("Outputs")
    for col, year in enumerate([1, 2, 3], start=2):
        engine.write(0, col - 1, year)
        outputs.write(0, col - 1, year)
        # Shared skeleton ``=<col>5*10`` so Pass-1 can parameterize by TIME_PERIOD.
        engine.write_number(4, col - 1, float(year))
        engine.write_formula(5, col - 1, f"={chr(64 + col)}5*10")
        outputs.write_formula(11, col - 1, f"=Engine!{chr(64 + col)}6")
    wb.close()


_COMPUTE_BINDINGS: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "compute_wire.xlsx",
    "series": [
        {
            "id": "engine_path",
            "sheet": "Engine",
            "data_range": "Engine!B6:D6",
            "layout": "series",
            "internal": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [_time_period_dim(header_row=1)],
            },
            "key": ["TIME_PERIOD"],
            "validation": {"intersect_graph_formulas": True},
        },
        {
            "id": "output_path",
            "sheet": "Outputs",
            "data_range": "Outputs!B12:D12",
            "layout": "series",
            "output": {
                "compute": {
                    "name": "compute_output_path",
                    "record_contract": "records",
                }
            },
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [_time_period_dim(header_row=1)],
            },
            "key": ["TIME_PERIOD"],
            "validation": {"intersect_graph_formulas": True},
        },
    ],
}


def test_compute_auto_wires_output_helper(tmp_path: Path) -> None:
    workbook = tmp_path / "compute_wire.xlsx"
    _write_compute_workbook(workbook)
    targets = expand_data_range("Outputs!B12:D12", workbook=workbook) + expand_data_range(
        "Engine!B6:D6", workbook=workbook
    )
    files = _generate(workbook, bindings_doc=_COMPUTE_BINDINGS, targets=targets)
    internals = files["internals.py"]
    api = files["api.py"]

    assert_helper_inventory(internals, {"output_path"})
    assert_helper_signature(internals, "output_path", ("ctx", "time_period"))
    assert_compute_calls_helper(
        api,
        "compute_output_path",
        "output_path",
        output_addresses=["Outputs!B12", "Outputs!C12", "Outputs!D12"],
    )


def _write_constant_leaf_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    consts = wb.add_worksheet("Consts")
    consts.write_number("A1", 1.5)
    engine = wb.add_worksheet("Engine")
    engine.write(0, 2, 1)
    engine.write_formula("C5", "=Consts!A1*2")
    wb.close()


_CONSTANT_LEAF_BINDINGS: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "constant_leaf.xlsx",
    "series": [
        {
            "id": "scale_factor",
            "sheet": "Consts",
            "data_range": "Consts!A1",
            "layout": "scalar",
            "constant": {},
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
        {
            "id": "scaled",
            "sheet": "Engine",
            "data_range": "Engine!C5",
            "layout": "scalar",
            "internal": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [],
            },
            "key": [],
            "validation": {"intersect_graph_formulas": True},
        },
    ],
}


def test_constant_leaf_rewritten_to_read(tmp_path: Path) -> None:
    workbook = tmp_path / "constant_leaf.xlsx"
    _write_constant_leaf_workbook(workbook)
    files = _generate(
        workbook,
        bindings_doc=_CONSTANT_LEAF_BINDINGS,
        targets=["Engine!C5", "Consts!A1"],
    )
    internals = files["internals.py"]
    assert_helper_inventory(internals, {"scaled"})
    body = _helper_body(internals, "scaled")
    assert "read_scale_factor" in body
    assert "xl_cell(ctx, 'Consts!A1')" not in body


def _write_mismatch_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    engine = wb.add_worksheet("Engine")
    engine.write(0, 2, 1)
    engine.write(0, 3, 2)
    engine.write_formula("C10", "=1+1")
    engine.write_formula("D10", "=SUM(100,200)")
    wb.close()


_MISMATCH_BINDINGS: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "mismatch.xlsx",
    "series": [
        {
            "id": "broken_family",
            "sheet": "Engine",
            "data_range": "Engine!C10:D10",
            "layout": "series",
            "internal": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [_time_period_dim(header_row=1)],
            },
            "key": ["TIME_PERIOD"],
            "validation": {"intersect_graph_formulas": True},
        },
    ],
}


def test_mixed_regime_leaves_intentional_cell_star_leftovers(tmp_path: Path) -> None:
    """Unverifiable clusters soft-skip: keep ``cell_*``, do not abort export.

    Mechanical synthesis failures (mixed-regime, non-parameterizable deps, …)
    leave intentional leftovers so downstream hybrid pipelines can finish those
    addresses locally.
    """
    workbook = tmp_path / "mismatch.xlsx"
    _write_mismatch_workbook(workbook)
    targets = expand_data_range("Engine!C10:D10", workbook=workbook)

    files = _generate(workbook, bindings_doc=_MISMATCH_BINDINGS, targets=targets)
    internals = files["internals.py"]
    cell_defs = sorted(n for n in def_names(internals) if n.startswith("cell_"))
    assert cell_defs == ["cell_engine_c10", "cell_engine_d10"]
    assert "mismatched" not in def_names(internals)
