"""Codegen emission contract for gating dependency tracking in exports (#238).

Non-iterative exports without input-direction series bindings omit ``deps`` /
``reverse_deps``, ``_record_dependency``, ``invalidate``, ``set_inputs``, and the
``ctx._record_dependency`` call site in ``_evaluate_address``.
Minimal non-iterative export baseline (``S!A1`` leaf + ``S!B1`` formula):

- **511 total lines**; embedded runtime **441 lines** (~86% of export)
- **0 dep-tracking lines**
- **62 cache-eval scaffold lines** (``_evaluate_address``, ``xl_cell``, ``xl_eval``)

Iterative exports and series bindings with input setters retain the full scaffold.
"""

from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any, cast

from excel_grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document
from tests.integration.user_flows.test_series_bindings_codegen import BINDINGS_DOCUMENT
from tests.integration.utils.parity_harness import (
    DEP_TRACKING_BASELINE_VERSION,
    SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET,
    assert_dep_tracking_absent,
    assert_dep_tracking_present,
    count_cache_eval_scaffold_lines,
    count_dep_tracking_lines,
    count_embedded_runtime_lines,
)
from tests.paths import DEP_TRACKING_BASELINE_FIXTURES

BASELINE_PATH = DEP_TRACKING_BASELINE_FIXTURES / "baseline.json"


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _minimal_non_iterative_graph() -> DependencyGraph:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 1.0))
    graph.add_node(_make_node("S!B1", "=S!A1+1", None))
    return graph


def _minimal_non_iterative_code() -> str:
    return CodeGenerator(_minimal_non_iterative_graph()).generate(["S!B1"])


def _load_baseline() -> dict[str, Any]:
    return cast(dict[str, Any], json.loads(BASELINE_PATH.read_text(encoding="utf-8")))


def test_dep_tracking_baseline_fixture_matches_schema() -> None:
    document = _load_baseline()
    assert document["version"] == DEP_TRACKING_BASELINE_VERSION
    metrics = document["minimal_non_iterative_export"]
    assert isinstance(metrics, dict)
    assert metrics["dep_tracking_lines"] == 0
    assert metrics["embedded_runtime_lines"] == 441
    targets = document["sprint2_targets"]
    assert targets["dep_tracking_lines"] == 0
    assert targets["cache_eval_scaffold_line_budget"] == SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET


def test_minimal_non_iterative_export_matches_baseline_metrics() -> None:
    """Record slim emission size for the canonical minimal export."""
    code = _minimal_non_iterative_code()
    baseline = _load_baseline()["minimal_non_iterative_export"]

    total_lines = len(code.splitlines())
    embedded_lines = count_embedded_runtime_lines(code)
    scaffold_lines = count_cache_eval_scaffold_lines(code)
    dep_lines = count_dep_tracking_lines(code)

    assert total_lines == baseline["total_export_lines"]
    assert embedded_lines == baseline["embedded_runtime_lines"]
    assert scaffold_lines == baseline["cache_eval_scaffold_lines"]
    assert dep_lines == baseline["dep_tracking_lines"]
    assert round(100 * embedded_lines / total_lines, 1) == baseline["embedded_runtime_share_pct"]


def test_minimal_non_iterative_export_omits_dep_tracking() -> None:
    """One-shot exports should not embed invalidation machinery."""
    assert_dep_tracking_absent(_minimal_non_iterative_code())


def test_minimal_non_iterative_export_meets_slim_scaffold_budget() -> None:
    code = _minimal_non_iterative_code()
    line_count = count_cache_eval_scaffold_lines(code)
    assert line_count <= SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET


def test_iterative_export_retains_dep_tracking() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", "=S!A1+1", None))
    code = CodeGenerator(graph, iterate_enabled=True).generate(["S!A1"])

    assert "xl_iterative_compute" in code
    assert_dep_tracking_present(code)


def test_series_binding_export_retains_dep_tracking_and_setters(
    tmp_path: Path,
) -> None:
    """Input setters call ``ctx.set_inputs``; export must keep invalidation scaffold."""
    from tests.integration.user_flows.utils import write_series_bindings_workbook

    workbook = tmp_path / "series_bindings.xlsx"
    write_series_bindings_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(BINDINGS_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)

    code = CodeGenerator(graph).generate(
        targets,
        series_bindings=bindings,
        bindings_workbook=workbook,
    )

    assert "def set_borvelia_primary_balance(" in code
    assert "ctx.set_inputs(" in code
    assert_dep_tracking_present(code)


def test_dep_tracking_gate_contract_documents_required_modes() -> None:
    """Inventory: slim vs full emission modes."""
    assert SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET == 62
