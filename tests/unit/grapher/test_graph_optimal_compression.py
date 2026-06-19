from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.core.cell_types import CellKind, EnumDomain
from excel_grapher.grapher.compression import OptimalCompressionRecord
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefLimits
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node


def _make_node(
    key: str,
    formula: str | None,
    normalized: str | None,
    *,
    is_leaf: bool = False,
    is_target: bool = False,
) -> Node:
    sheet, rest = key.split("!", 1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1]
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized,
        value=None,
        is_leaf=is_leaf,
        is_target=is_target,
    )


def _direct_edge(
    graph: DependencyGraph,
    dependent: str,
    precedent: str,
    *,
    formula: str | None = None,
    normalized: str | None = None,
) -> None:
    dr = DependencyCause.direct_ref
    dep_node = graph.get_node(dependent)
    assert dep_node is not None
    f = formula if formula is not None else dep_node.formula
    n = normalized if normalized is not None else dep_node.normalized_formula
    assert f is not None and n is not None
    ref = precedent
    i_f = f.index(ref)
    i_n = n.index(ref)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=((i_f, i_f + len(ref)),),
            direct_sites_normalized=((i_n, i_n + len(ref)),),
        ),
    )


def test_optimal_inline_single_call_site() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 5)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" in removed
    assert "Sheet1!B1" not in graph
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!D1"})
    na = graph.get_node("Sheet1!A1")
    assert na is not None
    assert na.normalized_formula == "=(Sheet1!D1*2)+1"


def test_optimal_inline_chain() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 2)
    c = _make_node("Sheet1!C1", "=Sheet1!D1+1", "=Sheet1!D1+1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1*2", "=Sheet1!C1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+3", "=Sheet1!B1+3")
    for n in (d, c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!C1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert {"Sheet1!B1", "Sheet1!C1"} <= set(removed)
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!D1"})
    na = graph.get_node("Sheet1!A1")
    assert na is not None
    assert na.normalized_formula is not None
    assert "Sheet1!D1" in na.normalized_formula
    assert "Sheet1!B1" not in na.normalized_formula
    assert "Sheet1!C1" not in na.normalized_formula


def test_optimal_subsumes_identity_transit() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 42)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" in removed
    assert "Sheet1!B1" not in graph
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!C1"})


def test_optimal_does_not_collapse_literal_leaf_formula() -> None:
    graph = DependencyGraph()
    b = _make_node("Sheet1!B1", "=1+1", "=1+1", is_leaf=True)
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    graph.add_node(b)
    graph.add_node(a)
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph


def test_optimal_preserve_blocks_inline() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2", is_target=True)
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph


def test_optimal_explicit_preserve_protects_public_cell() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal(preserve={"Sheet1!B1"})
    assert "Sheet1!B1" not in removed


def test_optimal_identity_replacement_protected_from_later_inline() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 1)
    c = _make_node("Sheet1!C1", "=Sheet1!D1", "=Sheet1!D1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1*2", "=Sheet1!C1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (d, c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!C1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!C1" in removed
    assert "Sheet1!C1" not in graph
    assert "Sheet1!D1" in graph


def test_optimal_keeps_multi_dependent_node() -> None:
    graph = DependencyGraph()
    g13 = _make_node("Sheet1!G13", "=AVERAGE(Sheet1!B1:Sheet1!B5)", "=AVERAGE(Sheet1!B1:Sheet1!B5)")
    g15 = _make_node(
        "Sheet1!G15", "=Sheet1!F15+(Sheet1!G13*Sheet1!G14)", "=Sheet1!F15+(Sheet1!G13*Sheet1!G14)"
    )
    g16 = _make_node("Sheet1!G16", "=Sheet1!G13+Sheet1!D16", "=Sheet1!G13+Sheet1!D16")
    for n in (g13, g15, g16):
        graph.add_node(n)
    for dep in ("Sheet1!F15", "Sheet1!G13", "Sheet1!G14"):
        graph.add_node(_make_node(dep, None, None, is_leaf=True))
    graph.add_node(_make_node("Sheet1!D16", None, None, is_leaf=True))
    _direct_edge(graph, "Sheet1!G15", "Sheet1!G13")
    _direct_edge(graph, "Sheet1!G16", "Sheet1!G13")

    removed = graph.compress_optimal()
    assert "Sheet1!G13" not in removed


def test_optimal_keeps_two_call_sites() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+Sheet1!B1", "=Sheet1!B1+Sheet1!B1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    f = "=Sheet1!B1+Sheet1!B1"
    ref = "Sheet1!B1"
    dr = DependencyCause.direct_ref
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=(
                (f.index(ref), f.index(ref) + len(ref)),
                (f.rindex(ref), f.rindex(ref) + len(ref)),
            ),
            direct_sites_normalized=(
                (f.index(ref), f.index(ref) + len(ref)),
                (f.rindex(ref), f.rindex(ref) + len(ref)),
            ),
        ),
    )

    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed


def test_optimal_blocks_unsafe_incoming_edge(tmp_path: Path) -> None:
    path = tmp_path / "rng.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 1)
    ws.write_number(0, 2, 2)
    ws.write_formula(0, 0, "=SUM(Sheet1!B1:C1)", None, 3)
    ws.write_formula(0, 3, "=Sheet1!B1*2", None, 2)
    ws.write_formula(0, 4, "=Sheet1!D1+1", None, 3)
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!E1"],
        load_values=False,
        capture_dependency_provenance=True,
    )
    removed = graph.compress_optimal()
    assert "Sheet1!D1" not in removed


def test_optimal_blocks_unsafe_body_range() -> None:
    graph = DependencyGraph()
    for key in ("Sheet1!B1", "Sheet1!B2", "Sheet1!B3"):
        graph.add_node(_make_node(key, None, None, is_leaf=True))
    g13 = _make_node("Sheet1!G13", "=AVERAGE(Sheet1!B1:Sheet1!B3)", "=AVERAGE(Sheet1!B1:Sheet1!B3)")
    g15 = _make_node("Sheet1!G15", "=Sheet1!G13+1", "=Sheet1!G13+1")
    graph.add_node(g13)
    graph.add_node(g15)
    graph.add_edge(
        "Sheet1!G13",
        "Sheet1!B1",
        provenance=EdgeProvenance(causes=frozenset({DependencyCause.static_range})),
    )
    _direct_edge(graph, "Sheet1!G15", "Sheet1!G13")

    removed = graph.compress_optimal()
    assert "Sheet1!G13" not in removed


def test_optimal_blocks_cyclic_candidate() -> None:
    graph = DependencyGraph()
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    b = _make_node("Sheet1!B1", "=Sheet1!A1*2", "=Sheet1!A1*2")
    graph.add_node(a)
    graph.add_node(b)
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")

    removed = graph.compress_optimal()
    assert removed == []


def test_optimal_does_not_mutate_original_when_projected_via_copy() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    original_keys = set(graph)
    projected = graph.copy()
    projected.compress_optimal()
    assert set(graph) == original_keys
    assert "Sheet1!B1" in graph


def test_optimal_record_captures_original_snapshot() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    record = OptimalCompressionRecord()
    graph.compress_optimal(record=record)
    assert record.snapshots_by_removed["Sheet1!B1"].formula == "=Sheet1!D1*2"
    assert record.inlined_to["Sheet1!B1"] == "Sheet1!A1"
    assert "Sheet1!B1" in record.removal_order


def test_optimal_blocks_dynamic_offset_body(tmp_path: Path) -> None:
    path = tmp_path / "off.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 2, 0)
    ws.write_formula(0, 1, "=OFFSET(Sheet1!C1,0,0)", None, 0)
    ws.write_formula(0, 0, "=Sheet1!B1+1", None, 1)
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!A1"],
        load_values=False,
        use_cached_dynamic_refs=True,
        capture_dependency_provenance=True,
    )
    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed


def test_optimal_blocks_when_provenance_absent() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    graph.add_edge("Sheet1!B1", "Sheet1!D1")
    graph.add_edge("Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed


@pytest.mark.parametrize("test_name", ["guard"])
def test_optimal_respects_guards(test_name: str) -> None:
    from excel_grapher.grapher.guard import Literal

    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!D1",
        guard=Literal(True),
        provenance=EdgeProvenance(causes=frozenset({dr})),
    )
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed


def test_optimal_blocks_indirect_enum_body(tmp_path: Path) -> None:
    path = tmp_path / "ind.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 5, 99)
    ws.write_formula(0, 0, "=Sheet1!F1", None, 99)
    ws.write_string(0, 3, "Sheet1!A1")
    ws.write_formula(0, 4, "=INDIRECT(Sheet1!D1)+Sheet1!A1", None, 0)
    ws.write_formula(0, 1, "=Sheet1!E1+1", None, 1)
    wb.close()

    from excel_grapher.core.cell_types import CellType

    env_typed = {
        "Sheet1!D1": CellType(
            kind=CellKind.STRING,
            enum=EnumDomain(values=frozenset({"Sheet1!A1:A10", "Sheet1!B1"})),
        )
    }
    cfg = DynamicRefConfig(
        cell_type_env=env_typed, limits=DynamicRefLimits(max_branches=16, max_cells=500)
    )

    graph = create_dependency_graph(
        path,
        ["Sheet1!B1"],
        load_values=False,
        dynamic_refs=cfg,
        capture_dependency_provenance=True,
    )
    removed = graph.compress_optimal()
    assert "Sheet1!E1" not in removed
