"""Opt-in raw formula storage and normalized-only provenance spans (issue #492).

`EdgeProvenance` no longer carries raw-formula spans; `normalized_formula` is the
single source for compression / projection rewriting, and `Node.formula` is an
audit-only field that extraction stores only when asked.
"""

from __future__ import annotations

import dataclasses
from pathlib import Path

import fastpyxl
import pytest

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.grapher.cache import (
    GRAPH_CACHE_SCHEMA_VERSION,
    _edge_provenance_from_json,
    _edge_provenance_to_json,
    dependency_graph_from_json,
    dependency_graph_to_json,
)
from excel_grapher.grapher.compression import FormulaRewrite
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.range_compression import (
    RawFormulasRequiredError,
    build_taco_index,
)


def _chain_workbook(path: Path) -> Path:
    """A1 -> B1 (identity transit) -> C1 (target), plus a second precedent."""
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 2
    ws["A2"].value = 3
    ws["B1"].value = "=A1"
    ws["C1"].value = "=B1*A2"
    wb.save(path)
    wb.close()
    return path


# ---- EdgeProvenance ------------------------------------------------------


def test_edge_provenance_has_no_raw_formula_spans() -> None:
    names = {f.name for f in dataclasses.fields(EdgeProvenance)}
    assert "direct_sites_formula" not in names
    assert "direct_sites_normalized" in names


def test_edge_provenance_rejects_raw_formula_spans_kwarg() -> None:
    # Passed dynamically so the removed keyword stays invisible to the type checker.
    legacy_kwargs: dict[str, object] = {"direct_sites_formula": ((1, 3),)}
    with pytest.raises(TypeError):
        EdgeProvenance(causes=DependencyCause.direct_ref, **legacy_kwargs)


def test_merge_unions_normalized_spans_only() -> None:
    a = EdgeProvenance(causes=DependencyCause.direct_ref, direct_sites_normalized=((1, 11),))
    b = EdgeProvenance(causes=DependencyCause.static_range, direct_sites_normalized=((20, 30),))
    merged = a.merge(b)
    assert merged.causes == DependencyCause.direct_ref | DependencyCause.static_range
    assert merged.direct_sites_normalized == ((1, 11), (20, 30))


# ---- cache serialization -------------------------------------------------


def test_cache_schema_version_bumped_past_raw_spans() -> None:
    assert GRAPH_CACHE_SCHEMA_VERSION >= 4


def test_provenance_json_omits_raw_formula_spans() -> None:
    blob = _edge_provenance_to_json(
        EdgeProvenance(causes=DependencyCause.direct_ref, direct_sites_normalized=((1, 11),))
    )
    assert "direct_sites_formula" not in blob
    assert blob["direct_sites_normalized"] == [[1, 11]]


def test_provenance_json_ignores_legacy_raw_formula_spans() -> None:
    restored = _edge_provenance_from_json(
        {
            "causes": ["direct_ref"],
            "direct_sites_formula": [[1, 3]],
            "direct_sites_normalized": [[1, 11]],
        }
    )
    assert restored.causes == DependencyCause.direct_ref
    assert restored.direct_sites_normalized == ((1, 11),)


def test_schema_3_graph_payload_still_loads(tmp_path: Path) -> None:
    """A serialized schema-3 graph carries raw-formula spans; loading must not fail."""
    graph = create_dependency_graph(
        _chain_workbook(tmp_path / "chain.xlsx"),
        ["Sheet1!C1"],
        capture_dependency_provenance=True,
    )
    payload = dependency_graph_to_json(graph)
    for edge in payload["edges"]:
        prov = edge["attrs"].get("provenance")
        if prov is not None:
            prov["direct_sites_formula"] = [[1, 3]]

    restored = dependency_graph_from_json(payload)
    assert restored.get_dependencies("Sheet1!C1") == {"Sheet1!B1", "Sheet1!A2"}
    prov = restored.get_edge_attrs("Sheet1!C1", "Sheet1!B1").provenance
    assert prov is not None
    assert prov.direct_sites_normalized == ((1, 1 + len("Sheet1!B1")),)


# ---- TACO range compression needs raw formulas ---------------------------


def test_build_taco_index_requires_raw_formulas(tmp_path: Path) -> None:
    graph = create_dependency_graph(_chain_workbook(tmp_path / "chain.xlsx"), ["Sheet1!C1"])
    with pytest.raises(RawFormulasRequiredError, match="store_raw_formula=True"):
        build_taco_index(graph)


def test_build_taco_index_accepts_graph_with_raw_formulas(tmp_path: Path) -> None:
    graph = create_dependency_graph(
        _chain_workbook(tmp_path / "chain.xlsx"),
        ["Sheet1!C1"],
        store_raw_formula=True,
    )
    index = build_taco_index(graph)
    materialized = {key for ref in index.find_precedents("Sheet1!C1") for key in ref.cell_keys()}
    assert materialized == {"Sheet1!B1", "Sheet1!A2"}


# ---- extraction opt-in ---------------------------------------------------


def test_raw_formula_is_not_stored_by_default(tmp_path: Path) -> None:
    graph = create_dependency_graph(_chain_workbook(tmp_path / "chain.xlsx"), ["Sheet1!C1"])
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.formula is None
    assert node.normalized_formula == "=Sheet1!B1*Sheet1!A2"
    assert node.is_leaf is False


def test_store_raw_formula_opt_in_keeps_workbook_text(tmp_path: Path) -> None:
    graph = create_dependency_graph(
        _chain_workbook(tmp_path / "chain.xlsx"),
        ["Sheet1!C1"],
        store_raw_formula=True,
    )
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.formula == "=B1*A2"
    assert node.normalized_formula == "=Sheet1!B1*Sheet1!A2"


def test_formula_node_iteration_ignores_raw_formula(tmp_path: Path) -> None:
    graph = create_dependency_graph(_chain_workbook(tmp_path / "chain.xlsx"), ["Sheet1!C1"])
    assert graph.formula_keys() == ["Sheet1!B1", "Sheet1!C1"]
    assert {k for k, _ in graph.formula_nodes()} == {"Sheet1!B1", "Sheet1!C1"}


def test_evaluator_runs_without_raw_formulas(tmp_path: Path) -> None:
    graph = create_dependency_graph(_chain_workbook(tmp_path / "chain.xlsx"), ["Sheet1!C1"])
    assert FormulaEvaluator(graph).evaluate("Sheet1!C1") == 6


# ---- provenance spans -----------------------------------------------------


def _qualified_workbook(path: Path) -> Path:
    """Formulas already sheet-qualified, so raw text equals normalized text."""
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 2
    ws["A2"].value = 3
    ws["B1"].value = "=Sheet1!A1"
    ws["C1"].value = "=Sheet1!B1*Sheet1!A2"
    wb.save(path)
    wb.close()
    return path


def test_spans_collected_when_raw_equals_normalized(tmp_path: Path) -> None:
    """An already-qualified formula takes the single-pass path; it still needs spans."""
    graph = create_dependency_graph(
        _qualified_workbook(tmp_path / "qualified.xlsx"),
        ["Sheet1!C1"],
        capture_dependency_provenance=True,
        store_raw_formula=True,
    )
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.formula == node.normalized_formula == "=Sheet1!B1*Sheet1!A2"

    prov = graph.get_edge_attrs("Sheet1!C1", "Sheet1!B1").provenance
    assert prov is not None
    assert prov.direct_sites_normalized == ((1, 1 + len("Sheet1!B1")),)


def test_branch_spans_are_not_offset_against_sub_expressions(tmp_path: Path) -> None:
    """Conditional branches are collected recursively; their spans must stay unrecorded.

    The recursion sees only the branch text, so any span it produced would index the
    sub-expression rather than the node's normalized formula.
    """
    path = tmp_path / "branches.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = 1
    ws["A2"].value = 10
    ws["A3"].value = 20
    ws["B1"].value = "=IF(Sheet1!A1=1,Sheet1!A2,Sheet1!A3)"
    wb.save(path)
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!B1"], capture_dependency_provenance=True)
    node = graph.get_node("Sheet1!B1")
    assert node is not None
    normalized = node.normalized_formula
    assert normalized is not None

    for precedent in ("Sheet1!A2", "Sheet1!A3"):
        prov = graph.get_edge_attrs("Sheet1!B1", precedent).provenance
        assert prov is not None
        for start, end in prov.direct_sites_normalized:
            assert normalized[start:end] == precedent


def test_compression_works_when_raw_equals_normalized(tmp_path: Path) -> None:
    graph = create_dependency_graph(
        _qualified_workbook(tmp_path / "qualified.xlsx"),
        ["Sheet1!C1"],
        capture_dependency_provenance=True,
    )
    assert "Sheet1!B1" in graph.compress_identity_transits()
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!A1*Sheet1!A2"


# ---- compression on normalized formulas only -----------------------------


def test_identity_transit_compression_without_raw_formulas(tmp_path: Path) -> None:
    graph = create_dependency_graph(
        _chain_workbook(tmp_path / "chain.xlsx"),
        ["Sheet1!C1"],
        capture_dependency_provenance=True,
    )
    removed = graph.compress_identity_transits()
    assert "Sheet1!B1" in removed
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.normalized_formula == "=Sheet1!A1*Sheet1!A2"
    assert FormulaEvaluator(graph).evaluate("Sheet1!C1") == 6


def test_formula_rewrite_records_normalized_formulas_only() -> None:
    names = {f.name for f in dataclasses.fields(FormulaRewrite)}
    assert names == {"dependent", "before_normalized", "after_normalized"}


def test_compression_leaves_raw_formula_as_captured(tmp_path: Path) -> None:
    """Raw formula is an audit record of workbook text; compression rewrites normalized."""
    graph = create_dependency_graph(
        _chain_workbook(tmp_path / "chain.xlsx"),
        ["Sheet1!C1"],
        capture_dependency_provenance=True,
        store_raw_formula=True,
    )
    graph.compress_identity_transits()
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    assert node.formula == "=B1*A2"
    assert node.normalized_formula == "=Sheet1!A1*Sheet1!A2"
