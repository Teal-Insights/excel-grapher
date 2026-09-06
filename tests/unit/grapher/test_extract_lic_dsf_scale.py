"""Regression for issue #716 (LIC-DSF-scale `create_dependency_graph`).

Ops counts are the oracle, not wall-clock. The tests encode the four extract
hot-path wins:

- Shape-keyed INDEX/OFFSET inference so row-wise copies share `_dyn_cache`.
- Nested-IF provenance is not a second IF-splitting walk.
- Copied formulas share one AST parse of the punched skeleton.
- Argument-env expansion is iterative (not a 400-frame Python recursion).
"""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest
import xlsxwriter

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    IntervalDomain,
)
from excel_grapher.core.formula_ast import (
    BinaryOpNode,
    CellRef,
    CellRefNode,
    RelativeAxis,
    clear_shape_parse_cache,
    parse_preserving_axes,
    shape_parse_cache_info,
)
from excel_grapher.grapher import dynamic_refs as dynamic_refs_mod
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.dependency_provenance import DependencyCause
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefConfig,
    DynamicRefLimits,
    DynamicRefTraceEvent,
    expand_leaf_env_to_argument_env,
    trace_dynamic_refs,
)


def _bfs_done(events: list[DynamicRefTraceEvent]) -> DynamicRefTraceEvent:
    done = [event for event in events if event.kind == "bfs-done"]
    assert done, f"expected bfs-done, got {[event.kind for event in events]}"
    return done[-1]


def test_row_wise_index_copies_share_shape_keyed_dyn_cache(tmp_path: Path) -> None:
    """Identical-shape INDEX rows over a fixed array must hit `_dyn_cache`.

    `INDEX($D$1:$D$20, MATCH(C{row}, $C$1:$C$20, 0))` copied down is the
    LIC-DSF pattern: lookup bases stay put, host A1 and MATCH lookup shift.
    Before shape-keyed inference, `trace_dynamic_refs` reported `infer_calls=N`
    and `cache_hits=0`.
    """
    n_rows = 20
    excel_path = tmp_path / "index_shape_cache.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(1, 21):
        ws.write_number(row - 1, 2, row)  # C1:C20 lookup keys
        ws.write_number(row - 1, 3, row * 10)  # D1:D20 values
    for row in range(1, n_rows + 1):
        ws.write_formula(
            row - 1,
            4,
            f"=INDEX($D$1:$D$20,MATCH(C{row},$C$1:$C$20,0))",
        )
    wb.close()

    env: CellTypeEnv = {}
    for row in range(1, 21):
        env[f"Sheet1!C{row}"] = CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=1, max=20),
        )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())
    targets = [f"Sheet1!E{row}" for row in range(1, n_rows + 1)]
    events: list[DynamicRefTraceEvent] = []
    with trace_dynamic_refs(events.append):
        graph = create_dependency_graph(
            excel_path,
            targets,
            load_values=False,
            dynamic_refs=config,
            capture_dependency_provenance=True,
        )

    stats = _bfs_done(events).detail
    infer_calls = stats["infer_calls"]
    cache_hits = stats["cache_hits"]
    assert infer_calls == 1, (
        f"expected one INDEX inference for {n_rows} identical-shape rows, "
        f"got infer_calls={infer_calls} cache_hits={cache_hits}"
    )
    assert cache_hits >= n_rows - 1, (
        f"row-wise INDEX copies must hit the shape-keyed cache; "
        f"infer_calls={infer_calls} cache_hits={cache_hits}"
    )

    first_index_deps = {
        dep for dep in graph.get_dependencies(targets[0]) if dep.startswith("Sheet1!D")
    }
    last_index_deps = {
        dep for dep in graph.get_dependencies(targets[-1]) if dep.startswith("Sheet1!D")
    }
    assert first_index_deps == last_index_deps
    assert "Sheet1!D1" in first_index_deps
    assert "Sheet1!D20" in first_index_deps
    assert "Sheet1!C1" in graph.get_dependencies(targets[0])
    assert f"Sheet1!C{n_rows}" in graph.get_dependencies(targets[-1])
    for target in targets:
        for dep in graph.get_dependencies(target):
            prov = graph.get_edge_attrs(target, dep).provenance
            assert prov is not None
            if dep.startswith("Sheet1!D"):
                assert DependencyCause.dynamic_index in prov.causes


def test_shifted_index_arrays_keep_per_cell_targets(tmp_path: Path) -> None:
    """When the INDEX array itself shifts, inferred targets must stay per-row."""
    excel_path = tmp_path / "shifted_index_array.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(1, 12):
        ws.write_number(row - 1, 0, row)
        ws.write_number(row - 1, 1, row * 10)
    for row in range(1, 6):
        ws.write_formula(
            row - 1,
            2,
            f"=INDEX(B{row}:B{row + 4},MATCH(A{row},A{row}:A{row + 4},0))",
        )
    wb.close()

    env: CellTypeEnv = {}
    for row in range(1, 12):
        env[f"Sheet1!A{row}"] = CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=1, max=12),
        )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())
    graph = create_dependency_graph(
        excel_path,
        [f"Sheet1!C{row}" for row in range(1, 6)],
        load_values=False,
        dynamic_refs=config,
        capture_dependency_provenance=True,
    )

    assert "Sheet1!B1" in graph.get_dependencies("Sheet1!C1")
    assert "Sheet1!B5" in graph.get_dependencies("Sheet1!C1")
    assert "Sheet1!B6" not in graph.get_dependencies("Sheet1!C1")
    assert "Sheet1!B5" in graph.get_dependencies("Sheet1!C5")
    assert "Sheet1!B9" in graph.get_dependencies("Sheet1!C5")
    assert "Sheet1!B1" not in graph.get_dependencies("Sheet1!C5")


def test_nested_if_copies_do_not_resplit_for_provenance(tmp_path: Path) -> None:
    """Provenance on nested IF copies must not re-enter the IF-splitting walk."""
    import excel_grapher.grapher.builder as builder_mod

    n_rows = 40
    excel_path = tmp_path / "nested_if_mcve.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Engine")
    ws.write_number(0, 0, 1)
    ws.write_number(0, 1, 2)
    ws.write_number(0, 2, 3)
    ws.write_number(0, 3, 4)
    for row in range(2, n_rows + 2):
        ws.write_formula(row - 1, 6, "=IF(A1>0,IF(B1>0,A1+B1,C1),IF(C1>0,D1,A1))")
    wb.close()

    original_collect = builder_mod.collect_provenance_for_formula
    collect_calls = 0

    def counting_collect(*args: object, **kwargs: object):
        nonlocal collect_calls
        collect_calls += 1
        return original_collect(*args, **kwargs)

    with patch.object(builder_mod, "collect_provenance_for_formula", side_effect=counting_collect):
        graph = create_dependency_graph(
            excel_path,
            [f"Engine!G{row}" for row in range(2, n_rows + 2)],
            load_values=False,
            capture_dependency_provenance=True,
        )

    assert collect_calls == 0, (
        "nested IF provenance should be accumulated during extract_deps_with_guards, "
        f"not via collect_provenance_for_formula ({collect_calls} calls)"
    )
    target = "Engine!G2"
    for dep in ("Engine!A1", "Engine!B1", "Engine!C1", "Engine!D1"):
        assert dep in graph.get_dependencies(target)
        prov = graph.get_edge_attrs(target, dep).provenance
        assert prov is not None
        assert DependencyCause.direct_ref in prov.causes


def test_copied_formulas_share_shape_keyed_parse(tmp_path: Path) -> None:
    """Relative copies must parse one punched skeleton, not one AST per A1 string."""
    n_rows = 30
    excel_path = tmp_path / "shape_parse.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(1, n_rows + 1):
        ws.write_number(row - 1, 0, row)
        ws.write_formula(row - 1, 1, f"=IF(A{row}>0,A{row}+A{row},A{row})")
    wb.close()

    clear_shape_parse_cache()
    graph = create_dependency_graph(
        excel_path,
        [f"Sheet1!B{row}" for row in range(1, n_rows + 1)],
        load_values=False,
    )
    hits, misses = shape_parse_cache_info()
    assert misses == 1, f"expected one skeleton parse, got misses={misses} hits={hits}"
    assert hits >= n_rows - 1, f"copied IF formulas must hit the shape parse cache; hits={hits}"
    b1 = graph._get_internal_node("Sheet1!B1")
    b2 = graph._get_internal_node("Sheet1!B2")
    assert b1 is not None and b2 is not None
    assert b1.formula_ast is b2.formula_ast


def test_parse_preserving_axes_shape_cache_matches_direct_parse() -> None:
    """Skeleton fill must preserve relative offsets of a direct parse."""
    clear_shape_parse_cache()
    first = parse_preserving_axes("=A1+B1", anchor="Sheet1!C3")
    second = parse_preserving_axes("=A2+B2", anchor="Sheet1!C4")
    assert first == second
    assert isinstance(first, BinaryOpNode)
    assert isinstance(first.left, CellRefNode)
    assert first.left.ref == CellRef(
        sheet="Sheet1",
        col=RelativeAxis(-2),
        row=RelativeAxis(-2),
    )
    hits, misses = shape_parse_cache_info()
    assert misses == 1
    assert hits == 1


def test_expand_env_is_iterative_not_recursive(monkeypatch: pytest.MonkeyPatch) -> None:
    """A 400-cell argument chain must not consume 400 Python stack frames."""
    depth = 400
    formulas = {f"Sheet1!A{row}": f"=Sheet1!A{row - 1}+1" for row in range(2, depth + 1)}
    refs = {f"=Sheet1!A{row - 1}+1": {f"Sheet1!A{row - 1}"} for row in range(2, depth + 1)}

    def get_cell_formula(addr: str) -> str | None:
        return formulas.get(addr)

    def get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        return set(refs.get(formula, set()))

    leaf_env: CellTypeEnv = {
        "Sheet1!A1": CellType(
            kind=CellKind.NUMBER,
            enum=EnumDomain(values=frozenset({1, 2})),
        )
    }
    monkeypatch.setattr(dynamic_refs_mod, "_MAX_ANALYSIS_DEPTH", depth + 10)
    previous_limit = sys.getrecursionlimit()
    sys.setrecursionlimit(200)
    try:
        env = expand_leaf_env_to_argument_env(
            {f"Sheet1!A{depth}"},
            get_cell_formula,
            get_refs_from_formula,
            leaf_env,
            DynamicRefLimits(),
        )
    finally:
        sys.setrecursionlimit(previous_limit)

    assert env[f"Sheet1!A{depth}"].enum is not None
    assert env["Sheet1!A1"].enum is not None
