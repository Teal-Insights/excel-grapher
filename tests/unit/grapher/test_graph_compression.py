from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.core.cell_types import CellKind, EnumDomain
from excel_grapher.grapher.compression import CompressionProvenanceRequiredError
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


def _direct_edge(graph: DependencyGraph, dependent: str, precedent: str) -> None:
    dep_node = graph.get_node(dependent)
    assert dep_node is not None
    normalized = dep_node.normalized_formula
    assert normalized is not None
    start = normalized.index(precedent)
    graph.add_edge(
        dependent,
        precedent,
        provenance=EdgeProvenance(
            causes=DependencyCause.direct_ref,
            direct_sites_normalized=((start, start + len(precedent)),),
        ),
    )


def _identity_transit_graph(*, b_is_target: bool = False) -> DependencyGraph:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 42)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1", is_target=b_is_target)
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for node in (c, b, a):
        graph.add_node(node)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    return graph


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
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=dr))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=dr,
            direct_sites_normalized=sp,
        ),
    )

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    assert "Sheet1!B1" not in graph
    assert graph.get_dependencies("Sheet1!A1") == {"Sheet1!C1"}
    na = graph.get_node("Sheet1!A1")
    assert na is not None
    assert na.normalized_formula == "=Sheet1!C1"


def test_compress_identity_transits_populates_record() -> None:
    from excel_grapher.grapher.compression import IdentityTransitCompressionRecord
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
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=dr))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=dr,
            direct_sites_normalized=sp,
        ),
    )

    record = IdentityTransitCompressionRecord()
    removed = graph.compress_identity_transits(record=record)

    assert removed == ["Sheet1!B1"]
    assert record.immediate_removed["Sheet1!B1"] == "Sheet1!C1"
    assert record.removal_order == ["Sheet1!B1"]
    assert len(record.formula_rewrites) == 1
    assert record.formula_rewrites[0].dependent == "Sheet1!A1"
    assert record.snapshots_by_removed["Sheet1!B1"].address == "Sheet1!B1"


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
    graph.add_edge("Sheet1!C1", "Sheet1!D1", provenance=EdgeProvenance(causes=dr))
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=dr))
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=dr,
            direct_sites_normalized=sp,
        ),
    )

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    assert "Sheet1!C1" in removed
    assert graph.get_dependencies("Sheet1!A1") == {"Sheet1!D1"}


def test_identity_transit_raises_when_provenance_missing(tmp_path: Path) -> None:
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
        graph.compress_identity_transits()


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
    assert "Sheet1!B1" in graph.get_dependencies("Sheet1!A1")
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
    prov = graph.get_edge_attrs("Sheet1!A1", "Sheet1!B1").provenance
    assert prov is not None
    assert DependencyCause.dynamic_offset in prov.causes
    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" not in removed


def test_index_blocks_compression(tmp_path: Path) -> None:
    path = tmp_path / "idx.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 2, 1)  # C1 row selector
    ws.write_number(0, 0, 10)  # A1 INDEX target when row=1
    ws.write_formula(0, 1, "=Sheet1!A1", None, 10)  # B1 transit of A1
    ws.write_formula(0, 3, "=INDEX(Sheet1!B1:Sheet1!B1,Sheet1!C1,1)", None, 10)  # D1
    wb.close()

    from excel_grapher.core.cell_types import CellType

    env = {
        "Sheet1!C1": CellType(
            kind=CellKind.NUMBER,
            enum=EnumDomain(values=frozenset({1})),
        )
    }
    cfg = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        path,
        ["Sheet1!D1"],
        load_values=False,
        dynamic_refs=cfg,
        capture_dependency_provenance=True,
    )
    prov = graph.get_edge_attrs("Sheet1!D1", "Sheet1!B1").provenance
    assert prov is not None
    assert DependencyCause.dynamic_index in prov.causes
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
    graph.add_edge("Sheet1!B1", "Sheet1!C1", provenance=EdgeProvenance(causes=dr))
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(causes=dr | dy),
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
        provenance=EdgeProvenance(causes=dr),
    )
    af = "=Sheet1!B1"
    ref = "Sheet1!B1"
    i = af.index(ref)
    sp = ((i, i + len(ref)),)
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(
            causes=dr,
            direct_sites_normalized=sp,
        ),
    )

    assert graph.compress_identity_transits() == []


def test_provenance_absent_raises_for_identity_transit() -> None:
    from excel_grapher.grapher.graph import DependencyGraph

    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for n in (c, b, a):
        graph.add_node(n)
    graph.add_edge("Sheet1!B1", "Sheet1!C1")
    graph.add_edge("Sheet1!A1", "Sheet1!B1")

    with pytest.raises(CompressionProvenanceRequiredError, match="provenance"):
        graph.compress_identity_transits()


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
    prov = graph.get_edge_attrs("Sheet1!E1", "Sheet1!A1").provenance
    assert prov is not None
    assert DependencyCause.direct_ref in prov.causes
    assert DependencyCause.dynamic_indirect in prov.causes
    assert graph.compress_identity_transits() == []


def test_identity_transit_is_target_not_removed() -> None:
    graph = _identity_transit_graph(b_is_target=True)

    removed = graph.compress_identity_transits()

    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})


def test_identity_transit_non_target_still_removed() -> None:
    graph = _identity_transit_graph()

    removed = graph.compress_identity_transits()

    assert "Sheet1!B1" in removed
    assert "Sheet1!B1" not in graph
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!C1"})


def test_identity_transit_explicit_preserve_not_removed() -> None:
    graph = _identity_transit_graph()

    removed = graph.compress_identity_transits(preserve={"Sheet1!B1"})

    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})


def test_identity_transit_is_target_protected_even_with_unrelated_preserve() -> None:
    graph = _identity_transit_graph(b_is_target=True)

    removed = graph.compress_identity_transits(preserve={"Sheet1!Z99"})

    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph


def test_identity_transit_target_barrier_does_not_block_upstream_collapse() -> None:
    graph = DependencyGraph()
    d = _make_node("Sheet1!D1", None, None, is_leaf=True)
    object.__setattr__(d, "value", 1)
    c = _make_node("Sheet1!C1", "=Sheet1!D1", "=Sheet1!D1")
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1", is_target=True)
    a = _make_node("Sheet1!A1", "=Sheet1!B1", "=Sheet1!B1")
    for node in (d, c, b, a):
        graph.add_node(node)
    _direct_edge(graph, "Sheet1!C1", "Sheet1!D1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")

    removed = graph.compress_identity_transits()

    assert "Sheet1!B1" not in removed
    assert "Sheet1!C1" in removed
    assert "Sheet1!B1" in graph
    assert "Sheet1!C1" not in graph
    b_node = graph.get_node("Sheet1!B1")
    assert b_node is not None
    assert b_node.normalized_formula == "=Sheet1!D1"
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!D1"})


def test_identity_transit_skips_when_whole_column_mentions_transit() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 1)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=SUM(Sheet1!B:B)", "=SUM(Sheet1!B:B)")
    for node in (c, b, a):
        graph.add_node(node)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph
    na = graph.get_node("Sheet1!A1")
    assert na is not None
    assert na.normalized_formula == "=SUM(Sheet1!B:B)"


def test_identity_transit_skips_when_whole_row_mentions_transit() -> None:
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 1)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=SUM(Sheet1!1:1)", "=SUM(Sheet1!1:1)")
    for node in (c, b, a):
        graph.add_node(node)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    graph.add_edge(
        "Sheet1!A1",
        "Sheet1!B1",
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )
    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" not in removed
    assert "Sheet1!B1" in graph
    na = graph.get_node("Sheet1!A1")
    assert na is not None
    assert na.normalized_formula == "=SUM(Sheet1!1:1)"


def test_identity_transit_skips_if_any_dependent_has_range_endpoint() -> None:
    """A second, cell-only dependent does not override a range-endpoint mention."""
    graph = DependencyGraph()
    c = _make_node("Sheet1!C1", None, None, is_leaf=True)
    object.__setattr__(c, "value", 1)
    b = _make_node("Sheet1!B1", "=Sheet1!C1", "=Sheet1!C1")
    a = _make_node("Sheet1!A1", "=SUM(Sheet1!B1:B3)", "=SUM(Sheet1!B1:B3)")
    d = _make_node("Sheet1!D1", "=Sheet1!B1", "=Sheet1!B1")
    for node in (c, b, a, d):
        graph.add_node(node)
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    _direct_edge(graph, "Sheet1!D1", "Sheet1!B1")

    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" not in removed
    assert graph.get_dependencies("Sheet1!D1") == {"Sheet1!B1"}
    nd = graph.get_node("Sheet1!D1")
    assert nd is not None
    assert nd.normalized_formula == "=Sheet1!B1"
