"""Tests for input.mode: override (non-leaf formula cell setters)."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

import xlsxwriter

from excel_grapher.core import CellValue
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict, xl_cell
from excel_grapher.series_bindings import derive_input_series, resolve_series_binding
from excel_grapher.series_bindings.setter_codegen import emit_setters_block
from excel_grapher.series_bindings.types import WorkbookSeriesBindings
from excel_grapher.series_bindings.validate import validate_series_bindings


def _write_override_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws_in = wb.add_worksheet("Inputs")
    ws_in.write_number("A1", 10)
    ws_eng = wb.add_worksheet("Engine")
    ws_eng.write_formula("B1", "=Inputs!A1+1")
    ws_out = wb.add_worksheet("Output")
    ws_out.write_formula("C1", "=Engine!B1*2")
    for col in range(2, 5):
        col_letter = chr(ord("A") + col - 1)
        ws_eng.write_number(f"{col_letter}1", col - 1)
        ws_eng.write_formula(f"{col_letter}2", f"=Inputs!A1+{col}")
        ws_out.write_formula(f"{col_letter}3", f"=Engine!{col_letter}2*2")
    wb.close()


def _override_scalar_series(*, series_id: str = "engine_override") -> dict[str, object]:
    return {
        "id": series_id,
        "sheet": "Engine",
        "data_range": "Engine!B1",
        "layout": "scalar",
        "input": {"mode": "override", "setter": {"name": "set_engine_b1"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "dtype": "float", "bind": {"kind": "data_cell"}},
            "dimensions": [],
        },
    }


def _manual_override_graph() -> DependencyGraph:
    graph = DependencyGraph()
    graph.add_node(
        Node(
            sheet="Inputs",
            column="A",
            row=1,
            formula=None,
            normalized_formula=None,
            value=10,
            is_leaf=True,
        )
    )
    graph.add_node(
        Node(
            sheet="Engine",
            column="B",
            row=1,
            formula="=Inputs!A1+1",
            normalized_formula="=Inputs!A1+1",
            value=11,
            is_leaf=False,
        )
    )
    graph.add_node(
        Node(
            sheet="Output",
            column="C",
            row=1,
            formula="=Engine!B1*2",
            normalized_formula="=Engine!B1*2",
            value=22,
            is_leaf=False,
            is_target=True,
        )
    )
    graph.add_edge("Engine!B1", "Inputs!A1")
    graph.add_edge("Output!C1", "Engine!B1")
    return graph


def test_resolve_input_override_includes_non_leaf_formula_cell(tmp_path: Path) -> None:
    wb_path = tmp_path / "override.xlsx"
    _write_override_workbook(wb_path)
    graph = _manual_override_graph()
    series = _override_scalar_series()

    resolved = resolve_series_binding(graph, wb_path, series, direction="input")

    engine_b1 = graph.get_node("Engine!B1")
    assert engine_b1 is not None
    assert engine_b1.is_leaf is False
    assert len(resolved["leaves"]) == 1
    assert resolved["leaves"][0]["address"] == "Engine!B1"
    assert not any(i["code"] == "partial_graph_overlap" for i in resolved["issues"])


def test_resolve_input_override_series_includes_formula_row(tmp_path: Path) -> None:
    wb_path = tmp_path / "override.xlsx"
    _write_override_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        ["Engine!B2", "Engine!C2", "Engine!D2", "Output!B3", "Output!C3", "Output!D3"],
        load_values=True,
    )
    series = {
        "id": "engine_row_override",
        "sheet": "Engine",
        "data_range": "Engine!B2:D2",
        "layout": "series",
        "input": {"mode": "override", "setter": {"name": "set_engine_row"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "dtype": "float", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "IDX",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                }
            ],
        },
        "key": ["IDX"],
    }

    resolved = resolve_series_binding(graph, wb_path, series, direction="input")

    assert [leaf["address"] for leaf in resolved["leaves"]] == [
        "Engine!B2",
        "Engine!C2",
        "Engine!D2",
    ]
    assert all(
        (node := graph.get_node(addr)) is not None and not node.is_leaf
        for addr in ["Engine!B2", "Engine!C2", "Engine!D2"]
    )


def _override_bindings() -> WorkbookSeriesBindings:
    return cast(
        WorkbookSeriesBindings,
        {"schema_version": "1.6.0", "series": [_override_scalar_series()]},
    )


def test_validate_leaf_mode_errors_on_non_leaf_overlap(tmp_path: Path) -> None:
    wb_path = tmp_path / "override.xlsx"
    _write_override_workbook(wb_path)
    graph = _manual_override_graph()
    series = _override_scalar_series()
    series = dict(series)
    series["input"] = {"setter": {"name": "set_engine_b1"}}

    report = validate_series_bindings(graph, {"schema_version": "1.6.0", "series": [series]})

    assert report["ok"] is False
    assert any(i["code"] == "non_leaf_input_overlap" for i in report["issues"])


def test_validate_override_requires_at_least_one_formula_node(tmp_path: Path) -> None:
    wb_path = tmp_path / "override.xlsx"
    _write_override_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!A1"], load_values=True)
    series = {
        "id": "leaf_only_override",
        "sheet": "Inputs",
        "data_range": "Inputs!A1",
        "layout": "scalar",
        "input": {"mode": "override", "setter": {"name": "set_inputs_a1"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "dtype": "float", "bind": {"kind": "data_cell"}},
            "dimensions": [],
        },
    }

    report = validate_series_bindings(graph, {"schema_version": "1.6.0", "series": [series]})

    assert report["ok"] is False
    assert any(i["code"] == "no_formula_override_targets" for i in report["issues"])


def test_validate_override_mode_accepts_mixed_leaf_and_formula_cells(tmp_path: Path) -> None:
    wb_path = tmp_path / "mixed.xlsx"
    wb = xlsxwriter.Workbook(wb_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("A1", 1)
    ws.write_formula("B1", "=A1+1")
    wb.close()

    graph = create_dependency_graph(wb_path, ["Sheet1!A1", "Sheet1!B1"], load_values=True)
    series = {
        "id": "mixed_override",
        "sheet": "Sheet1",
        "data_range": "Sheet1!A1:B1",
        "layout": "series",
        "input": {"mode": "override", "setter": {"name": "set_mixed"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "dtype": "float", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "IDX",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "read": "int"},
                }
            ],
        },
        "key": ["IDX"],
    }

    report = validate_series_bindings(graph, {"schema_version": "1.6.0", "series": [series]})

    assert report["ok"] is True


def test_derive_input_series_from_override_binding(tmp_path: Path) -> None:
    wb_path = tmp_path / "override.xlsx"
    _write_override_workbook(wb_path)
    graph = _manual_override_graph()
    bindings = _override_bindings()

    input_series = derive_input_series(graph, bindings, workbook=wb_path)

    assert len(input_series) == 1
    assert input_series[0]["setter_name"] == "set_engine_b1"
    assert [cell["address"] for cell in input_series[0]["cells"]] == ["Engine!B1"]


def test_override_setter_invalidates_parent_cache(tmp_path: Path) -> None:
    wb_path = tmp_path / "override.xlsx"
    _write_override_workbook(wb_path)
    graph = _manual_override_graph()
    bindings = _override_bindings()

    def make_resolver() -> Callable[[str], Callable[[EvalContext], CellValue] | None]:
        def _engine_b1(ctx: EvalContext) -> CellValue:
            return cast(CellValue, cast(float, xl_cell(ctx, "Inputs!A1")) + 1)

        def _output_c1(ctx: EvalContext) -> CellValue:
            return cast(CellValue, cast(float, xl_cell(ctx, "Engine!B1")) * 2)

        impls: dict[str, Callable[[EvalContext], CellValue]] = {
            "Engine!B1": _engine_b1,
            "Output!C1": _output_c1,
        }
        return lambda addr: impls.get(addr)

    lines = emit_setters_block(graph, wb_path, bindings, include_helpers=True)
    ns: dict[str, object] = {"EvalContext": EvalContext, "coerce_inputs_dict": coerce_inputs_dict}
    exec("\n".join(lines), ns)

    resolver = make_resolver()
    ctx = EvalContext(inputs=coerce_inputs_dict({"Inputs!A1": 10}), resolver=resolver)
    assert xl_cell(ctx, "Output!C1") == 22

    setter = cast(Callable[[EvalContext, object], None], ns["set_engine_b1"])
    setter(ctx, 99)

    assert ctx.inputs["Engine!B1"] == 99
    assert "Engine!B1" not in ctx.cache
    assert "Output!C1" not in ctx.cache
    assert xl_cell(ctx, "Output!C1") == 198
