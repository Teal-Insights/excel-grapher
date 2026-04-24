from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.core.cell_types import CellKind, EnumDomain
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefLimits
from excel_grapher.grapher.node import Node


def _make_node(
    key: str,
    formula: str | None,
    normalized: str | None,
    *,
    is_leaf: bool = False,
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
    )


def test_compress_happy_path_manual_graph() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = Node(
        sheet="Sheet1",
        column="C",
        row=1,
        formula=None,
        normalized_formula=None,
        value=42,
        is_leaf=True,
    )
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    graph.add_node(c)
    graph.add_node(b)
    graph.add_node(a)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    assert "Sheet1!B1" not in graph
    assert graph.dependencies("Sheet1!A1") == {"Sheet1!C1"}
    na = graph.get_node("Sheet1!A1")
    assert na is not None
    assert na.normalized_formula == "=Sheet1!C1"


def test_compress_chain_manual_graph() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 1)
    c = _make_node("Sheet1!C1", "=Sheet1!D1", "=Sheet1!D1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (d, c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge("Sheet1!C1", "Sheet1!D1", provenance=EdgeProvenance(causes=frozenset({dr})))
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    assert "Sheet1!C1" in removed
    assert graph.dependencies("Sheet1!A1") == {"Sheet1!D1"}


def test_static_range_blocks_compression(tmp_path: Path) -> None:
    path = tmp_path / "rng.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 1)
    ws.write_number(0, 2, 2)
    ws.write_formula(0, 0, "=SUM(Sheet1!B1:C1)", None, 3)
    ws.write_formula(0, 3, "=Sheet1!B1", None, 1)
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!A1"],
        load_values=False,
        capture_dependency_provenance=True,
    )
    assert "Sheet1!B1" in graph.dependencies("Sheet1!A1")
    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph


def test_offset_blocks_compression(tmp_path: Path) -> None:
    path = tmp_path / "off.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 2, 0)  # C1
    ws.write_formula(0, 1, "=Sheet1!C1", None, 0)  # B1 transit
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,0,0)", None, 0)  # A1
    wb.close()

    graph = create_dependency_graph(
        path,
        ["Sheet1!A1"],
        load_values=False,
        use_cached_dynamic_refs=True,
        capture_dependency_provenance=True,
    )
    prov = graph.edge_attrs("Sheet1!A1", "Sheet1!B1").get("provenance")
    assert prov is not None
    assert DependencyCause.dynamic_offset in prov.causes
    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" not in removed


def test_mixed_direct_and_offset_blocks_manual() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1+1", "=Sheet1!B1+1")
    for n in (c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    dy = DependencyCause.dynamic_offset
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=frozenset({dr})))
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(causes=frozenset({dr, dy})),
    )
    assert graph.compress_identity_transits() == []


def test_guarded_transit_not_compressed() -> None:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.guard import Literal

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    dr = DependencyCause.direct_ref
    graph.add_edge(
        "Sheet1!B1",
        "Sheet1!C1",
        guard=Literal(True),
        provenance=EdgeProvenance(causes=frozenset({dr})),
    )
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=sp,
            direct_sites_normalized=sp,
        ),
    )

    assert graph.compress_identity_transits() == []


def test_provenance_absent_skips_compression() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    graph.add_edge("Sheet1!B1", "Sheet1!C1")
    graph.add_edge("Sheet1!A1", "Sheet1!B1")

    assert graph.compress_identity_transits() == []


def test_indirect_enum_blocks_when_direct_same_cell(tmp_path: Path) -> None:
    path = tmp_path / "ind.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 5, 99)  # F1 leaf
    ws.write_formula(0, 0, "=Sheet1!F1", None, 99)  # A1 transit
    ws.write_string(0, 3, "Sheet1!A1")
    ws.write_formula(0, 4, "=INDIRECT(Sheet1!D1)+Sheet1!A1", None, 0)  # E1
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
        ["Sheet1!E1"],
        load_values=False,
        dynamic_refs=cfg,
        capture_dependency_provenance=True,
    )
    prov = graph.edge_attrs("Sheet1!E1", "Sheet1!A1").get("provenance")
    assert prov is not None
    assert DependencyCause.direct_ref in prov.causes
    assert DependencyCause.dynamic_indirect in prov.causes
    assert graph.compress_identity_transits() == []
