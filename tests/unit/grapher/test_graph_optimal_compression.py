from __future__ import annotations

from pathlib import Path
from typing import Annotated, Literal

import fastpyxl
import pytest
import xlsxwriter

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.cell_types import Between, CellKind, EnumDomain, RealBetween
from excel_grapher.evaluator.parser import parse
from excel_grapher.grapher.compression import (
    CompressionProvenanceRequiredError,
    OptimalCompressionRecord,
)
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
    normalized: str | None = None,
) -> None:
    dr = DependencyCause.direct_ref
    dep_node = graph.get_node(dependent)
    assert dep_node is not None
    n = normalized if normalized is not None else dep_node.normalized_formula
    assert n is not None
    ref = precedent
    i_n = n.index(ref)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=dr,
            direct_sites_normalized=((i_n, i_n + len(ref)),),
        ),
    )


def _local_ref_edge(
    graph: DependencyGraph,
    dependent: str,
    precedent: str,
    *,
    formula_ref: str,
    normalized_ref: str | None = None,
) -> None:
    dr = DependencyCause.direct_ref
    dep_node = graph.get_node(dependent)
    assert dep_node is not None
    f = dep_node.formula
    n = dep_node.normalized_formula
    assert f is not None and n is not None
    assert formula_ref in f
    ref_n = normalized_ref if normalized_ref is not None else precedent
    i_n = n.index(ref_n)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=dr,
            direct_sites_normalized=((i_n, i_n + len(ref_n)),),
        ),
    )


def _build_issue_277_workbook(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws_inputs = wb.active
    ws_inputs.title = "Inputs"
    ws_engine = wb.create_sheet("Engine")

    cells = {
        ("Inputs", "B6"): 100,
        ("Inputs", "B21"): 1,
        ("Inputs", "B22"): 1,
        ("Inputs", "C16"): 2.0,
        ("Inputs", "C17"): 3.0,
        ("Inputs", "C18"): -1.0,
        ("Engine", "B9"): 0.5,
        ("Engine", "C5"): 1,
        ("Engine", "B20"): "=Inputs!B6",
        ("Engine", "C10"): "=IF(C5>=Inputs!$B$21,1,0)",
        ("Engine", "C14"): "=Inputs!C16+CHOOSE(Inputs!$B$22,$B$9,0,0)*C10",
        ("Engine", "C15"): "=Inputs!C17+CHOOSE(Inputs!$B$22,0,$B$9,0)*C10",
        ("Engine", "C16"): "=Inputs!C18+CHOOSE(Inputs!$B$22,0,0,$B$9)*C10",
        ("Engine", "C20"): "=B20*(1+C15/100)/(1+C14/100)-C16",
    }
    for (sheet, addr), value in cells.items():
        worksheet = ws_inputs if sheet == "Inputs" else ws_engine
        worksheet[addr] = value
    wb.save(path)


def test_compress_identity_transit_refreshes_sibling_edge_spans() -> None:
    graph = DependencyGraph()
    b20 = _make_node("Engine!B20", "=Inputs!B6", "=Inputs!B6")
    c14 = _make_node(
        "Engine!C14",
        "=Inputs!C16+CHOOSE(Inputs!$B$22,$B$9,0,0)*C10",
        "=Inputs!C16+CHOOSE(Inputs!$B$22,Engine!$B$9,0,0)*Engine!C10",
    )
    c15 = _make_node("Engine!C15", "=Inputs!C17", "=Inputs!C17")
    c16 = _make_node("Engine!C16", "=Inputs!C18", "=Inputs!C18")
    c20 = _make_node(
        "Engine!C20",
        "=B20*(1+C15/100)/(1+C14/100)-C16",
        "=Engine!B20*(1+Engine!C15/100)/(1+Engine!C14/100)-Engine!C16",
    )
    for n in (b20, c14, c15, c16, c20):
        graph.add_node(n)
    graph.add_edge(
        "Engine!B20",
        "Inputs!B6",
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )
    _local_ref_edge(graph, "Engine!C20", "Engine!B20", formula_ref="B20")
    _local_ref_edge(graph, "Engine!C20", "Engine!C14", formula_ref="C14")
    _local_ref_edge(graph, "Engine!C20", "Engine!C15", formula_ref="C15")
    _local_ref_edge(graph, "Engine!C20", "Engine!C16", formula_ref="C16")

    graph._compress_one_transit("Engine!B20", "Inputs!B6")

    prov = graph.get_edge_attrs("Engine!C20", "Engine!C14").provenance
    assert prov is not None
    node = graph.get_node("Engine!C20")
    assert node is not None
    normalized = node.normalized_formula
    assert normalized is not None
    assert (
        normalized[prov.direct_sites_normalized[0][0] : prov.direct_sites_normalized[0][1]]
        == "Engine!C14"
    )
    # Compression rewrites normalized text only; the raw formula stays as captured.
    assert node.formula == "=B20*(1+C15/100)/(1+C14/100)-C16"


def test_optimal_inline_local_ref_after_identity_transit(tmp_path: Path) -> None:
    path = tmp_path / "issue_277.xlsx"
    _build_issue_277_workbook(path)

    constraints = {
        "Engine!C5": Literal[1],
        "Inputs!B21": Annotated[int, Between(1, 5)],
        "Inputs!B22": Literal[1, 2, 3],
        "Inputs!C16": Annotated[float, RealBetween(-10.0, 15.0)],
        "Inputs!C17": Annotated[float, RealBetween(0.0, 20.0)],
        "Inputs!C18": Annotated[float, RealBetween(-15.0, 15.0)],
    }
    config = DynamicRefConfig.from_constraints(constraints, {})
    graph = create_dependency_graph(
        path,
        ["Engine!C20"],
        load_values=True,
        dynamic_refs=config,
        capture_dependency_provenance=True,
    )

    projected = graph.copy()
    removed = projected.compress_optimal()
    assert "Engine!B20" in removed
    assert "Engine!C14" in removed
    # Embedded CHOOSE branch guard survives structural inline onto C20.
    from excel_grapher.grapher.guard import CellRef, Compare
    from excel_grapher.grapher.guard import Literal as GuardLiteral

    assert projected.get_edge_guard("Engine!C20", "Engine!B9") == Compare(
        left=CellRef(key="Inputs!B22"), op="=", right=GuardLiteral(value=1)
    )

    node = projected.get_node("Engine!C20")
    assert node is not None
    formula = node.normalized_formula
    assert formula is not None
    parse(formula.strip())
    assert "/100" in formula
    assert "(Inputs!C16+CHOOSE" in formula


def test_optimal_raises_when_provenance_missing(tmp_path: Path) -> None:
    path = tmp_path / "no_prov.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws_inputs = wb.add_worksheet("Inputs")
    ws_engine = wb.add_worksheet("Engine")
    ws_inputs.write_number(5, 1, 100)
    ws_engine.write_formula(19, 1, "=Inputs!B6", None, 100)
    ws_engine.write_formula(19, 2, "=B20*2", None, 200)
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Engine!C20"],
        load_values=False,
        capture_dependency_provenance=False,
    )
    with pytest.raises(CompressionProvenanceRequiredError, match="provenance"):
        graph.compress_optimal()


def test_optimal_no_raise_when_nothing_compressible(tmp_path: Path) -> None:
    path = tmp_path / "leaf.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!A1"],
        load_values=False,
        capture_dependency_provenance=False,
    )
    assert graph.compress_optimal() == []


def test_optimal_manual_graph_with_explicit_provenance_still_works() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" in removed


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


def test_optimal_inlines_transit_with_guarded_outgoing_edges() -> None:
    """Embedded conditionals leave guards on transit outs; the body is still pasteable."""
    from excel_grapher.grapher.guard import CellRef, Compare, Literal

    branch_guard = Compare(left=CellRef(key="Sheet1!C1"), op="=", right=Literal(value=1))
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node(
        "Sheet1!B1",
        "=IF(Sheet1!C1=1,Sheet1!D1,0)",
        "=IF(Sheet1!C1=1,Sheet1!D1,0)",
    )
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (c, d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!D1",
        guard=branch_guard,
        provenance=EdgeProvenance(
            causes=DependencyCause.direct_ref,
            direct_sites_normalized=((20, 29),),
        ),
    )
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" in removed
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!C1") is None
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!D1") == branch_guard
    na = graph.get_node("Sheet1!A1")
    assert na is not None
    assert na.normalized_formula == "=(IF(Sheet1!C1=1,Sheet1!D1,0))+1"


def test_optimal_inline_shared_dep_unguarded_on_dependent_wins() -> None:
    """Unconditional read on the dependent clears a guarded read from the transit."""
    from excel_grapher.grapher.guard import CellRef, Compare, Literal

    branch_guard = Compare(left=CellRef(key="Sheet1!C1"), op="=", right=Literal(value=1))
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node(
        "Sheet1!B1",
        "=IF(Sheet1!C1=1,Sheet1!D1,0)",
        "=IF(Sheet1!C1=1,Sheet1!D1,0)",
    )
    a = _make_node(
        "Sheet1!A1",
        "=Sheet1!B1+Sheet1!D1",
        "=Sheet1!B1+Sheet1!D1",
    )
    for n in (c, d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!D1",
        guard=branch_guard,
        provenance=EdgeProvenance(
            causes=DependencyCause.direct_ref,
            direct_sites_normalized=((20, 29),),
        ),
    )
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!D1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" in removed
    assert graph.get_edge_guard("Sheet1!A1", "Sheet1!D1") is None


def test_merge_inline_edge_guards_none_wins_and_or() -> None:
    from excel_grapher.grapher.compression import merge_inline_edge_guards
    from excel_grapher.grapher.guard import CellRef, Compare, Literal, Or

    g1 = Compare(left=CellRef(key="Sheet1!C1"), op="=", right=Literal(value=1))
    g2 = Compare(left=CellRef(key="Sheet1!C1"), op="=", right=Literal(value=2))

    assert (
        merge_inline_edge_guards(
            dependent_guard=None,
            dependent_has_edge=True,
            transit_guard=g1,
            transit_has_edge=True,
        )
        is None
    )
    assert (
        merge_inline_edge_guards(
            dependent_guard=g1,
            dependent_has_edge=True,
            transit_guard=None,
            transit_has_edge=True,
        )
        is None
    )
    assert (
        merge_inline_edge_guards(
            dependent_guard=g1,
            dependent_has_edge=True,
            transit_guard=g1,
            transit_has_edge=True,
        )
        == g1
    )
    assert merge_inline_edge_guards(
        dependent_guard=g1,
        dependent_has_edge=True,
        transit_guard=g2,
        transit_has_edge=True,
    ) == Or(operands=(g1, g2))
    assert (
        merge_inline_edge_guards(
            dependent_guard=None,
            dependent_has_edge=False,
            transit_guard=g1,
            transit_has_edge=True,
        )
        == g1
    )


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


def test_optimal_preserve_blocks_identity_transit() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 42)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1", is_target=True)
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})


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


def test_optimal_explicit_preserve_blocks_identity_transit() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 42)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal(preserve={"Sheet1!B1"})
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph


def test_optimal_is_target_protected_even_with_unrelated_preserve() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 42)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1", is_target=True)
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_optimal(preserve={"Sheet1!Z99"})
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph


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
            causes=dr,
            direct_sites_normalized=(
                (f.index(ref), f.index(ref) + len(ref)),
                (f.rindex(ref), f.rindex(ref) + len(ref)),
            ),
        ),
    )

    removed = graph.compress_optimal()
    assert "Sheet1!B1" not in removed


def test_optimal_blocks_unsafe_incoming_edge(tmp_path: Path) -> None:
    """A dependent reaching the transit through a range ref cannot absorb its body."""
    path = tmp_path / "rng.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 1)  # B1
    ws.write_formula(0, 3, "=Sheet1!B1*2", None, 2)  # D1
    ws.write_number(1, 3, 5)  # D2
    ws.write_formula(0, 4, "=SUM(Sheet1!D1:D2)", None, 7)  # E1

    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!E1"],
        load_values=False,
        capture_dependency_provenance=True,
    )
    prov = graph.get_edge_attrs("Sheet1!E1", "Sheet1!D1").provenance
    assert prov is not None
    assert DependencyCause.static_range in prov.causes

    removed = graph.compress_optimal()
    assert "Sheet1!D1" not in removed


def test_optimal_inlines_already_qualified_formulas(tmp_path: Path) -> None:
    """Normalization is a no-op on sheet-qualified formulas; spans must still be recorded."""
    path = tmp_path / "qualified.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 1)  # B1
    ws.write_formula(0, 3, "=Sheet1!B1*2", None, 2)  # D1
    ws.write_formula(0, 4, "=Sheet1!D1+1", None, 3)  # E1
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!E1"],
        load_values=False,
        capture_dependency_provenance=True,
    )
    before = FormulaEvaluator(graph).evaluate("Sheet1!E1")

    assert "Sheet1!D1" in graph.compress_optimal()
    node = graph.get_node("Sheet1!E1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!B1*2+1"
    assert graph.get_dependencies("Sheet1!E1") == {"Sheet1!B1"}
    assert FormulaEvaluator(graph).evaluate("Sheet1!E1") == before


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
        provenance=EdgeProvenance(causes=DependencyCause.static_range),
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


def test_optimal_blocks_guarded_incoming_edge() -> None:
    """A guarded dependent→transit edge is not inlined (guards are not conjoined)."""
    from excel_grapher.grapher.guard import Literal

    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!D1*2", "=Sheet1!D1*2")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (d, b, a):
        graph.add_node(n)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!D1")
    af = "=Sheet1!B1+1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        guard=Literal(True),
        provenance=EdgeProvenance(
            causes=DependencyCause.direct_ref,
            direct_sites_normalized=sp,
        ),
    )

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
