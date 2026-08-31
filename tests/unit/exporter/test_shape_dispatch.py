"""Shape-dispatch analysis and codegen: one function per interned formula shape."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import CodeGenerator, create_dependency_graph
from excel_grapher.core.formula_shape import fingerprint_formula_shape
from excel_grapher.exporter.shape_dispatch import (
    ConstantArg,
    GeometricCellArg,
    LeafHolePlan,
    LookupHolePlan,
    PassthroughHolePlan,
    analyze_shape_dispatch,
)
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


def _dag_workbook(path: Path) -> None:
    """Write a small DAG with several interned shapes (no cycles).

    Layout::

        A1:A3   leaves 10, 20, 30
        B1:B3   =A{row}+1                 shape add1
        C1:C3   =B{row}*2                 shape mul2 (passthrough → add1)
        D1      =SUM(A1:A3)               shape sum
        E1:E3   =C{row}+$D$1              shape add (passthrough + constant)
        G2      =$A$3+1                   add1, absolute (breaks geometry for L)
        L1      =B1*3                     shape mul3
        L2      =G2*3                     mul3 lookup (same child shape, irregular)
        N1      =C1-D1                    shape sub
        N2      =C2-B2                    sub lookup (mixed child shapes on p1)
    """
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30
    ws["B1"] = "=A1+1"
    ws["B2"] = "=A2+1"
    ws["B3"] = "=A3+1"
    ws["C1"] = "=B1*2"
    ws["C2"] = "=B2*2"
    ws["C3"] = "=B3*2"
    ws["D1"] = "=SUM(A1:A3)"
    ws["E1"] = "=C1+$D$1"
    ws["E2"] = "=C2+$D$1"
    ws["E3"] = "=C3+$D$1"
    ws["G2"] = "=$A$3+1"
    ws["L1"] = "=B1*3"
    ws["L2"] = "=G2*3"
    ws["N1"] = "=C1-D1"
    ws["N2"] = "=C2-B2"
    wb.save(path)
    wb.close()


def _dag_graph(path: Path):
    _dag_workbook(path)
    return create_dependency_graph(
        path,
        ["Sheet1!E1:E3", "Sheet1!L1", "Sheet1!L2", "Sheet1!N1", "Sheet1!N2"],
        load_values=True,
        warm_formula_shapes=True,
    )


def _shape_key(formula: str) -> str:
    return fingerprint_formula_shape(formula).shape_key


def _helper_names(graph) -> dict[str, str]:
    table = graph.formula_shapes
    assert table is not None
    keys = sorted({table.lookup(addr)[0] for addr in graph.formula_keys() if table.lookup(addr)})
    eligible = [key for key in keys if CodeGenerator._shape_helper_eligible(table.shapes[key])]
    return {key: f"_shape_{index}" for index, key in enumerate(eligible)}


def _plan_for(layout, shape_key: str):
    for plan in layout.plans:
        if plan.shape_key == shape_key:
            return plan
    raise AssertionError(f"no plan for {shape_key}")


def test_analyze_classifies_passthrough_constant_and_lookup(tmp_path: Path) -> None:
    graph = _dag_graph(tmp_path / "dag.xlsx")
    helpers = _helper_names(graph)
    layout = analyze_shape_dispatch(graph, list(graph.formula_keys()), helpers)

    add1 = _plan_for(layout, _shape_key("=Sheet1!A1+1"))
    assert isinstance(add1.holes[0], LeafHolePlan)
    assert set(add1.hosts) == {"Sheet1!B1", "Sheet1!B2", "Sheet1!B3", "Sheet1!G2"}

    mul2 = _plan_for(layout, _shape_key("=Sheet1!B1*2"))
    hole = mul2.holes[0]
    assert isinstance(hole, PassthroughHolePlan)
    assert hole.child_shape_key == add1.shape_key
    assert hole.args == (GeometricCellArg(dcol=-1, drow=0),)

    add = _plan_for(layout, _shape_key("=Sheet1!C1+Sheet1!D1"))
    p0, p1 = add.holes
    assert isinstance(p0, PassthroughHolePlan)
    assert p0.child_shape_key == mul2.shape_key
    assert p0.args == (GeometricCellArg(dcol=-1, drow=0),)
    assert isinstance(p1, PassthroughHolePlan)
    assert p1.args == (ConstantArg("Sheet1!A1:A3"),)

    mul3 = _plan_for(layout, _shape_key("=Sheet1!B1*3"))
    lookup = mul3.holes[0]
    assert isinstance(lookup, LookupHolePlan)
    hosts = {entry.host: entry.child_params for entry in lookup.entries}
    assert hosts["Sheet1!B1"] == ("Sheet1!A1",)
    assert hosts["Sheet1!G2"] == ("Sheet1!A3",)

    sub = _plan_for(layout, _shape_key("=Sheet1!C1-Sheet1!D1"))
    assert isinstance(sub.holes[0], PassthroughHolePlan)
    mixed = sub.holes[1]
    assert isinstance(mixed, LookupHolePlan)
    by_host = {entry.host: entry.child_shape_key for entry in mixed.entries}
    assert by_host["Sheet1!D1"] == _plan_for(layout, _shape_key("=SUM(Sheet1!A1:A3)")).shape_key
    assert by_host["Sheet1!B2"] == add1.shape_key


def test_shape_dispatch_emits_one_function_per_shape(tmp_path: Path) -> None:
    graph = _dag_graph(tmp_path / "dag.xlsx")
    code = CodeGenerator(graph, shape_dispatch=True).generate(
        ["Sheet1!E1", "Sheet1!E3", "Sheet1!L1", "Sheet1!L2", "Sheet1!N1", "Sheet1!N2"]
    )
    start = code.index("# --- Formula")
    end = code.index("# --- Formula resolver")
    section = code[start:end]
    assert "def cell_sheet1_" not in section
    assert section.count("def _shape_") == 6
    assert "_CELL_SHAPES" in section
    assert "_offset_cell(" in section
    assert "_eval_shape(" in section
    assert "_eval_lookup(" in section
    assert "_eval_shape(ctx, p0," in section
    assert "Sheet1!A1:A3" in section


def test_shape_dispatch_matches_evaluator(tmp_path: Path) -> None:
    graph = _dag_graph(tmp_path / "dag.xlsx")
    targets = [
        "Sheet1!E1",
        "Sheet1!E2",
        "Sheet1!E3",
        "Sheet1!L1",
        "Sheet1!L2",
        "Sheet1!N1",
        "Sheet1!N2",
    ]
    result = assert_codegen_matches_evaluator(
        graph,
        targets,
        shape_dispatch=True,
    )
    assert result.generated_results["Sheet1!E1"] == 82
    assert result.generated_results["Sheet1!E2"] == 102
    assert result.generated_results["Sheet1!E3"] == 122
    assert result.generated_results["Sheet1!L1"] == 33
    assert result.generated_results["Sheet1!L2"] == 93
    assert result.generated_results["Sheet1!N1"] == -38
    assert result.generated_results["Sheet1!N2"] == 21


def test_shape_dispatch_requires_warm_shapes(tmp_path: Path) -> None:
    _dag_workbook(tmp_path / "dag.xlsx")
    graph = create_dependency_graph(
        tmp_path / "dag.xlsx",
        ["Sheet1!E1"],
        load_values=True,
    )
    with pytest.raises(ValueError, match="formula_shapes"):
        CodeGenerator(graph, shape_dispatch=True).generate(["Sheet1!E1"])


def test_default_generate_still_emits_per_cell_wrappers(tmp_path: Path) -> None:
    graph = _dag_graph(tmp_path / "dag.xlsx")
    code = CodeGenerator(graph).generate(["Sheet1!C1", "Sheet1!C2"])
    assert "def cell_sheet1_c1" in code
    assert "def _shape_" in code
