"""Shared argument-subgraph ref-walk caches (issue #539)."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated

import fastpyxl

from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause
from excel_grapher.grapher.dynamic_ref_walk import DynamicRefWalkContext
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.grapher.parser import FormulaNormalizer


def _build_index_match_chain_workbook(path: Path) -> None:
    """Write the 20-row INDEX/MATCH grid from issue #539."""
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for row in range(1, 21):
        ws[f"A{row}"] = row
        ws[f"B{row}"] = row * 10
    for row in range(1, 21):
        ws[f"C{row}"] = f"=B{row}+A{row}"
    for row in range(1, 21):
        ws[f"D{row}"] = f"=C{row}*2"
    for row in range(1, 21):
        ws[f"E{row}"] = f"=INDEX(D1:D20,MATCH(C{row},C1:C20,0))"
    ws["F1"] = "=SUM(E1:E20)"
    wb.save(path)


def _index_match_dynamic_refs() -> DynamicRefConfig:
    schema = {
        **{f"Sheet1!A{row}": Annotated[int, "leaf"] for row in range(1, 21)},
        **{f"Sheet1!B{row}": Annotated[int, "leaf"] for row in range(1, 21)},
    }
    return DynamicRefConfig.from_constraints(schema, {})


def _count_parse_range_refs_with_spans(
    path: Path,
    *,
    capture_dependency_provenance: bool,
) -> int:
    from unittest.mock import patch

    from excel_grapher.grapher import parser as parser_mod

    orig = parser_mod.parse_range_refs_with_spans
    count = {"n": 0}

    def counted(formula: str) -> list:
        count["n"] += 1
        return orig(formula)

    with patch.object(parser_mod, "parse_range_refs_with_spans", counted):
        create_dependency_graph(
            path,
            ["Sheet1!F1"],
            dynamic_refs=_index_match_dynamic_refs(),
            capture_dependency_provenance=capture_dependency_provenance,
            load_values=False,
        )
    return count["n"]


def test_argument_subgraph_refs_returns_independent_sets() -> None:
    """Caller mutations of subgraph sets must not poison later lookups."""
    formulas = {("S", "C1"): "=A1+B1"}

    def get_cell_value(sheet: str, a1: str) -> object:
        return formulas.get((sheet, a1), 1)

    ctx = DynamicRefWalkContext(
        normalizer=FormulaNormalizer({}, {}),
        max_range_cells=5000,
        get_cell_value=get_cell_value,
        sheet_names={"S"},
    )
    all_refs, _leaves = ctx.argument_subgraph_refs({"S!C1"})
    all_refs.add("poison")
    all_refs_again, leaves_again = ctx.argument_subgraph_refs({"S!C1"})
    assert "poison" not in all_refs_again
    assert all_refs_again == {"S!C1", "S!A1", "S!B1"}
    assert leaves_again == {"S!A1", "S!B1"}


def test_refs_in_formula_without_dynamic_is_memoized() -> None:
    formulas = {("S", "C1"): "=A1+B1"}
    reads = {"n": 0}

    def get_cell_value(sheet: str, a1: str) -> object:
        reads["n"] += 1
        return formulas.get((sheet, a1), 1)

    ctx = DynamicRefWalkContext(
        normalizer=FormulaNormalizer({}, {}),
        max_range_cells=5000,
        get_cell_value=get_cell_value,
        sheet_names={"S"},
    )
    first = ctx.refs_in_formula_without_dynamic("=A1+B1", "S")
    second = ctx.refs_in_formula_without_dynamic("=A1+B1", "S")
    first.add("poison")
    assert second == {"S!A1", "S!B1"}
    ctx.argument_subgraph_refs({"S!C1"})
    ctx.argument_subgraph_refs({"S!C1"})
    # One worksheet read for C1; A1/B1 reads happen once via the node cache.
    assert reads["n"] == 3


def test_provenance_does_not_multiply_range_parses_on_index_match_chain(
    tmp_path: Path,
) -> None:
    """Provenance must reuse extract's cached argument-subgraph walk (issue #539).

    Before the shared `DynamicRefWalkContext`, enabling provenance caused ~9×
    more `parse_range_refs_with_spans` calls than extract-only on this grid.
    """
    path = tmp_path / "index_chain.xlsx"
    _build_index_match_chain_workbook(path)

    without = _count_parse_range_refs_with_spans(path, capture_dependency_provenance=False)
    with_prov = _count_parse_range_refs_with_spans(path, capture_dependency_provenance=True)

    assert without > 0
    ratio = with_prov / without
    # Issue #539 reproduced ~9× more parses with provenance enabled. Shared
    # walk caches plus skipping BFS on a `_dyn_cache` hit should keep the
    # remaining overhead (per-formula span collection) well below that.
    assert ratio < 3, (
        f"parse_range_refs_with_spans without provenance: {without}; "
        f"with provenance: {with_prov}; ratio: {ratio:.2f} (issue #539)"
    )


def test_index_match_chain_still_records_dynamic_index_provenance(tmp_path: Path) -> None:
    """Sharing the walk cache must not drop INDEX target provenance."""
    path = tmp_path / "index_chain_prov.xlsx"
    _build_index_match_chain_workbook(path)
    graph = create_dependency_graph(
        path,
        ["Sheet1!F1"],
        dynamic_refs=_index_match_dynamic_refs(),
        capture_dependency_provenance=True,
        load_values=False,
    )
    prov = graph.get_edge_attrs("Sheet1!E1", "Sheet1!D1").provenance
    assert prov is not None
    assert DependencyCause.dynamic_index in prov.causes
    prov_match = graph.get_edge_attrs("Sheet1!E1", "Sheet1!C1").provenance
    assert prov_match is not None
    assert DependencyCause.dynamic_index in prov_match.causes
