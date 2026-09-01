"""Pass-1 upstream gaps from Tiny DSA adoption feedback (#595).

Covers: collapse vs projection graph alignment, xl_memoize in runtime.py,
``_ADDRESS_DISPATCH`` in the resolver, and ``skip_collapse`` / ``clustering_mode``
knobs on ``generate_modules``.
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest
import xlsxwriter

from excel_grapher.exporter import CodeGenerator, OptimalCompression
from excel_grapher.exporter.pass1.models import SeriesHelperVerificationError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document
from tests.integration.exporter.pass1_shape_contract import (
    assert_helper_inventory,
    assert_no_cell_defs_for_addresses,
    def_names,
)


def _write_inline_chain_workbook(path: Path) -> None:
    """Engine hop + output: OptimalCompression can inline Engine!C6 into Outputs!B12."""
    wb = xlsxwriter.Workbook(path)
    engine = wb.add_worksheet("Engine")
    engine.write_number("B1", 10)
    engine.write_formula("C6", "=Engine!B1*2")
    out = wb.add_worksheet("Outputs")
    out.write(0, 1, 1)  # header year
    out.write_formula("B12", "=Engine!C6")
    out.write_formula("B14", "=Outputs!B12+1")
    wb.close()


_INLINE_BINDINGS: dict[str, Any] = {
    "schema_version": "1.9.0",
    "workbook": "inline.xlsx",
    "series": [
        {
            "id": "seed",
            "sheet": "Engine",
            "data_range": "Engine!B1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_seed"}},
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
            "id": "engine_hop",
            "sheet": "Engine",
            "data_range": "Engine!C6",
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
        {
            "id": "baseline",
            "sheet": "Outputs",
            "data_range": "Outputs!B12",
            "layout": "scalar",
            "output": {"compute": {"name": "compute_baseline"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "concept": "LABEL",
                        "role": "key",
                        "scope": "series",
                        "bind": {"kind": "constant", "value": "baseline"},
                    }
                ],
            },
            "key": ["LABEL"],
        },
        {
            "id": "delta",
            "sheet": "Outputs",
            "data_range": "Outputs!B14",
            "layout": "scalar",
            "output": {"compute": {"name": "compute_delta"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "concept": "LABEL",
                        "role": "key",
                        "scope": "series",
                        "bind": {"kind": "constant", "value": "delta"},
                    }
                ],
            },
            "key": ["LABEL"],
        },
    ],
}


def test_projected_public_only_preserve_fails_closed_with_guidance(tmp_path: Path) -> None:
    """Collapse must not cluster on original while IR comes from a thinner projection."""
    workbook = tmp_path / "inline.xlsx"
    _write_inline_chain_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(_INLINE_BINDINGS))
    targets = ["Outputs!B12", "Outputs!B14"]
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        capture_dependency_provenance=True,
    )

    # Preserve only public output addresses — engine_hop can be inlined away.
    from excel_grapher.series_bindings.workflow import series_binding_public_addresses

    public_only = frozenset(
        addr
        for addr in series_binding_public_addresses(graph, bindings, workbook=workbook)
        if addr.startswith("Outputs!")
    )
    projection = OptimalCompression(preserve=set(public_only) | set(targets)).project(graph)
    assert "Engine!C6" not in projection

    with (
        CodeGenerator(projection) as gen,
        pytest.raises(SeriesHelperVerificationError, match="preserve all bound series"),
    ):
        gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def test_projected_preserve_all_bound_series_collapses(tmp_path: Path) -> None:
    workbook = tmp_path / "inline.xlsx"
    _write_inline_chain_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(_INLINE_BINDINGS))
    targets = ["Outputs!B12", "Outputs!B14", "Engine!C6", "Engine!B1"]
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        capture_dependency_provenance=True,
    )
    from excel_grapher.series_bindings.workflow import series_binding_public_addresses

    preserve = set(series_binding_public_addresses(graph, bindings, workbook=workbook))
    projection = OptimalCompression(preserve=preserve).project(graph)
    assert "Engine!C6" in projection

    with CodeGenerator(projection) as gen:
        files = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )
    internals = files["internals.py"]
    assert_helper_inventory(internals, {"engine_hop", "baseline", "delta"})
    assert_no_cell_defs_for_addresses(internals, ["Engine!C6", "Outputs!B12", "Outputs!B14"])


def test_recurrence_embeds_xl_memoize_in_runtime(tmp_path: Path) -> None:
    from tests.integration.exporter.test_pass1_shape_units import (
        _RECURRENCE_BINDINGS,
        _write_recurrence_workbook,
    )

    workbook = tmp_path / "recurrence.xlsx"
    _write_recurrence_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(_RECURRENCE_BINDINGS))
    targets = expand_data_range("Engine!C6:E6", workbook=workbook) + [
        "Inputs!B1",
        "Inputs!B2",
    ]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        files = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )
    assert "@xl_memoize" in files["internals.py"]
    assert "def xl_memoize(" in files["runtime.py"]
    assert (
        "xl_memoize"
        in files["internals.py"].split("from .runtime import", 1)[1].split("\n\n", 1)[0]
    )


def test_row_series_emits_address_dispatch(tmp_path: Path) -> None:
    from tests.integration.exporter.test_pass1_shape_units import (
        _ROW_SERIES_BINDINGS,
        _write_row_series_workbook,
    )

    workbook = tmp_path / "row_series.xlsx"
    _write_row_series_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(_ROW_SERIES_BINDINGS))
    targets = expand_data_range("Engine!C10:G10", workbook=workbook) + ["Inputs!B1"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        files = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )
    internals = files["internals.py"]
    assert "_ADDRESS_DISPATCH = {" in internals
    assert "dispatch = _ADDRESS_DISPATCH.get(address)" in internals
    ns: dict[str, Any] = {}
    start = internals.index("_ADDRESS_DISPATCH = ")
    end = internals.index("\ndef _address_to_func_name", start)
    exec(internals[start:end], ns)
    dispatch = ns["_ADDRESS_DISPATCH"]
    assert dispatch["Engine!C10"] == ("shock_flag", {"time_period": 1})
    assert dispatch["Engine!G10"] == ("shock_flag", {"time_period": 5})


def test_mixed_regime_soft_skip_records_skipped_and_keeps_other_helpers(
    tmp_path: Path,
) -> None:
    """Partial collapse: uniform series still helperize; mixed-regime stays ``cell_*``."""
    from excel_grapher.exporter.pass1.collapse import collapse_bound_series_in_source
    from tests.integration.exporter.test_pass1_shape_units import _ROW_SERIES_BINDINGS

    workbook = tmp_path / "combined.xlsx"
    wb = xlsxwriter.Workbook(workbook)
    engine = wb.add_worksheet("Engine")
    for col, year in enumerate([1, 2, 3, 4, 5], start=3):
        engine.write(0, col - 1, year)
        engine.write_number(4, col - 1, float(year))
        engine.write_formula(9, col - 1, f"=IF(Engine!{chr(64 + col)}5>=Inputs!$B$1,1,0)")
    engine.write_formula("C11", "=1+1")
    engine.write_formula("D11", "=SUM(100,200)")
    inputs = wb.add_worksheet("Inputs")
    inputs.write_number("B1", 3)
    wb.close()

    bindings_doc = deepcopy(_ROW_SERIES_BINDINGS)
    bindings_doc["workbook"] = "combined.xlsx"
    bindings_doc["series"].append(
        {
            "id": "mismatched",
            "sheet": "Engine",
            "data_range": "Engine!C11:D11",
            "layout": "series",
            "internal": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {
                            "kind": "column_header",
                            "header_row": 1,
                            "read": "int",
                        },
                    }
                ],
            },
            "key": ["TIME_PERIOD"],
            "validation": {"intersect_graph_formulas": True},
        }
    )
    bindings = validate_bindings_document(bindings_doc)
    targets = (
        expand_data_range("Engine!C10:G10", workbook=workbook)
        + expand_data_range("Engine!C11:D11", workbook=workbook)
        + ["Inputs!B1"]
    )
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        files = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            skip_collapse=True,
        )
    collapse = collapse_bound_series_in_source(
        files["internals.py"],
        graph=graph,
        bindings=bindings,
        workbook=workbook,
    )
    assert "shock_flag" in collapse.helper_names
    assert collapse.skipped
    assert all(s.reason.startswith("mixed_regime_groups:") for s in collapse.skipped)
    remaining = def_names(collapse.source)
    assert "cell_engine_c11" in remaining
    assert "cell_engine_d11" in remaining

    with CodeGenerator(graph) as gen:
        files = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )
    names = def_names(files["internals.py"])
    assert "cell_engine_c11" in names
    assert "cell_engine_d11" in names
    assert_helper_inventory(files["internals.py"], {"shock_flag"})


def test_skip_collapse_keeps_cell_star_ir(tmp_path: Path) -> None:
    from tests.integration.exporter.test_pass1_shape_units import (
        _ROW_SERIES_BINDINGS,
        _write_row_series_workbook,
    )

    workbook = tmp_path / "row_series.xlsx"
    _write_row_series_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(_ROW_SERIES_BINDINGS))
    targets = expand_data_range("Engine!C10:G10", workbook=workbook) + ["Inputs!B1"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        files = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            skip_collapse=True,
        )
    internals = files["internals.py"]
    names = def_names(internals)
    assert "shock_active" not in names
    assert any(n.startswith("cell_engine_") for n in names)
    assert "_ADDRESS_DISPATCH" not in internals


def test_clustering_mode_ast_is_accepted(tmp_path: Path) -> None:
    """Knob is plumbed; series-blind AST mode may still emit helpers for this MCVE."""
    from tests.integration.exporter.test_pass1_shape_units import (
        _ROW_SERIES_BINDINGS,
        _write_row_series_workbook,
    )

    workbook = tmp_path / "row_series.xlsx"
    _write_row_series_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(_ROW_SERIES_BINDINGS))
    targets = expand_data_range("Engine!C10:G10", workbook=workbook) + ["Inputs!B1"]
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        files = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
            clustering_mode="series",
        )
    assert "def shock_flag(" in files["internals.py"]
