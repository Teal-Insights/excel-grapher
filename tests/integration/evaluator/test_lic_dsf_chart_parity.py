"""LIC-DSF Chart Data: evaluator matches Excel **cached** values on chart strips (integration, slow).

Builds graphs with `use_cached_dynamic_refs=True` and asserts `FormulaEvaluator`
matches last-saved workbook numbers for figure and stress ranges (evaluator ↔ Excel).
Live recalculation is covered elsewhere; strict dynamic-ref resolution without cache
remains tracked separately via `xfail` (issue #255).
"""

from __future__ import annotations

import pytest

from excel_grapher import DynamicRefError, create_dependency_graph
from tests.integration.evaluator.utils.lic_dsf_chart_targets import (
    GRAPH_MAX_DEPTH,
    WORKBOOK_PATH,
    chart_parity_shortlist_keys,
    collect_chart_data_cell_keys,
)
from tests.utils.excel_workbook_parity import (
    assert_workbook_parity,
    compare_evaluator_to_excel_cache,
    format_workbook_parity_report,
)

RTOL = 1e-5
ATOL = 1e-9


@pytest.fixture(scope="module")
def lic_dsf_chart_shortlist_graph():
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")
    targets = chart_parity_shortlist_keys()
    return create_dependency_graph(
        WORKBOOK_PATH,
        targets,
        load_values=True,
        max_depth=GRAPH_MAX_DEPTH,
        use_cached_dynamic_refs=True,
    )


@pytest.mark.slow
def test_lic_dsf_chart_shortlist_graph_has_numeric_excel_cache(
    lic_dsf_chart_shortlist_graph,
) -> None:
    """Graph build loads Excel cached numeric results for the chart parity probe cells."""
    for key in chart_parity_shortlist_keys():
        node = lic_dsf_chart_shortlist_graph.get_node(key)
        assert node is not None, f"missing node {key}"
        assert isinstance(node.value, (int, float)) and not isinstance(node.value, bool), (
            f"{key}: expected numeric cached value, got {node.value!r}"
        )


@pytest.mark.slow
def test_lic_dsf_chart_shortlist_evaluator_matches_excel_cache(
    lic_dsf_chart_shortlist_graph,
) -> None:
    """Cached path: MX shock (U63) and Threshold (U66) must match Excel cache on the graph nodes."""
    assert_workbook_parity(
        lic_dsf_chart_shortlist_graph,
        chart_parity_shortlist_keys(),
        rtol=RTOL,
        atol=ATOL,
        fail_fast=True,
    )


@pytest.mark.slow
@pytest.mark.xfail(
    raises=DynamicRefError,
    strict=True,
    reason="Issue #255: strict dynamic-ref resolution for the chart slice is incomplete.",
)
def test_lic_dsf_chart_shortlist_without_cached_dynamic_refs_raises() -> None:
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")
    create_dependency_graph(
        WORKBOOK_PATH,
        chart_parity_shortlist_keys(),
        load_values=True,
        max_depth=GRAPH_MAX_DEPTH,
        use_cached_dynamic_refs=False,
    )


@pytest.fixture(scope="module")
def lic_dsf_full_chart_graph():
    if not WORKBOOK_PATH.exists():
        pytest.skip(f"Test workbook not found at {WORKBOOK_PATH}")
    keys = collect_chart_data_cell_keys()
    return create_dependency_graph(
        WORKBOOK_PATH,
        keys,
        load_values=True,
        max_depth=GRAPH_MAX_DEPTH,
        use_cached_dynamic_refs=True,
    )


@pytest.mark.slow
def test_lic_dsf_full_chart_export_graph_builds(lic_dsf_full_chart_graph) -> None:
    """Full chart export targets produce a non-empty graph with cached dynamic refs."""
    assert len(lic_dsf_full_chart_graph) > 0
    export_keys = set(collect_chart_data_cell_keys())
    hits = sum(1 for k in lic_dsf_full_chart_graph if k in export_keys)
    assert hits == len(export_keys)


@pytest.mark.slow
def test_lic_dsf_full_chart_export_evaluator_matches_excel_cache(lic_dsf_full_chart_graph) -> None:
    """Cached path: formula cells in the export set with numeric Excel cache must match evaluator.

    Scope: intersection of export keys with formula nodes in the built graph.
    """
    export_keys = set(collect_chart_data_cell_keys())
    graph = lic_dsf_full_chart_graph
    candidates: list[str] = []
    for addr in graph.formula_keys():
        if addr not in export_keys:
            continue
        node = graph.get_node(addr)
        if node is None or not node.normalized_formula:
            continue
        if isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            candidates.append(addr)

    mismatches = compare_evaluator_to_excel_cache(
        graph, candidates, rtol=RTOL, atol=ATOL, fail_fast=False
    )
    if mismatches:
        by_kind: dict[str, int] = {}
        for m in mismatches:
            by_kind[m.kind.value] = by_kind.get(m.kind.value, 0) + 1
        summary = f"counts_by_kind={by_kind!r}\n"
        raise AssertionError(
            summary
            + format_workbook_parity_report(mismatches[:50])
            + (f"\n... and {len(mismatches) - 50} more" if len(mismatches) > 50 else "")
        )
