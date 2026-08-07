"""Tests for list_dynamic_ref_constraint_candidates."""

from __future__ import annotations

from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    IntervalDomain,
)
from excel_grapher.grapher.builder import list_dynamic_ref_constraint_candidates
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefLimits

# ---------------------------------------------------------------------------
# Workbook factories
# ---------------------------------------------------------------------------


def _build_no_dynamic_refs(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(1, 0, 10)  # A2
    ws.write_number(2, 0, 20)  # A3
    ws.write_formula(0, 0, "=A2+A3", None, 30)  # A1
    wb.close()


def _build_single_offset_missing_leaf(path: Path) -> None:
    """A1 = OFFSET(Sheet1!B1, 0, Sheet1!C1).  C1 is a leaf with no constraint."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(0, 2, 1)  # C1 row-offset leaf (missing constraint)
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,0,Sheet1!C1)", None, 10)  # A1
    wb.close()


def _build_two_offsets_missing_leaves(path: Path) -> None:
    """Build a workbook with two OFFSET formulas and missing leaf constraints.

    A1 = B1 + C1  (static formula)
    B1 = OFFSET(Sheet1!D1, 0, Sheet1!E1)   E1 missing
    C1 = OFFSET(Sheet1!F1, 0, Sheet1!G1)   G1 missing.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(0, 0, "=Sheet1!B1+Sheet1!C1", None, 0)  # A1
    ws.write_formula(0, 1, "=OFFSET(Sheet1!D1,0,Sheet1!E1)", None, 0)  # B1
    ws.write_formula(0, 2, "=OFFSET(Sheet1!F1,0,Sheet1!G1)", None, 0)  # C1
    ws.write_number(0, 3, 10)  # D1 base
    ws.write_number(0, 4, 1)  # E1 missing constraint
    ws.write_number(0, 5, 20)  # F1 base
    ws.write_number(0, 6, 1)  # G1 missing constraint
    wb.close()


def _build_all_leaves_constrained(path: Path) -> None:
    """A1 = OFFSET(Sheet1!B1, 0, Sheet1!C1).  C1 is constrained."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(0, 2, 1)  # C1 leaf (will be constrained)
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,0,Sheet1!C1)", None, 10)  # A1
    wb.close()


def _build_partial_constraint(path: Path) -> None:
    """Build a workbook with one constrained and one missing OFFSET leaf.

    A1 = OFFSET(Sheet1!B1, 0, Sheet1!C1) + OFFSET(Sheet1!B1, 0, Sheet1!D1)
    C1 is constrained, D1 is missing.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(0, 2, 0)  # C1 constrained
    ws.write_number(0, 3, 1)  # D1 missing
    ws.write_formula(
        0,
        0,
        "=OFFSET(Sheet1!B1,0,Sheet1!C1)+OFFSET(Sheet1!B1,0,Sheet1!D1)",
        None,
        10,
    )  # A1
    wb.close()


def _build_static_index_only(path: Path) -> None:
    """A1 = INDEX(Sheet1!B1:Sheet1!B3, 1, 1).  Static INDEX: no candidates."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1
    ws.write_number(1, 1, 20)  # B2
    ws.write_number(2, 1, 30)  # B3
    ws.write_formula(0, 0, "=INDEX(Sheet1!B1:Sheet1!B3,1,1)", None, 10)  # A1
    wb.close()


def _build_index_match_range_arg(path: Path) -> None:
    """D5 = INDEX(A10:C12, MATCH(B5, A10:A12, 0), 2)."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(4, 1, 20)  # B5 lookup value
    ws.write_number(9, 0, 10)  # A10
    ws.write_number(10, 0, 20)  # A11
    ws.write_number(11, 0, 30)  # A12
    ws.write_number(9, 1, 100)  # B10
    ws.write_number(10, 1, 200)  # B11
    ws.write_number(11, 1, 300)  # B12
    ws.write_number(9, 2, 1000)  # C10
    ws.write_number(10, 2, 2000)  # C11
    ws.write_number(11, 2, 3000)  # C12
    ws.write_formula(4, 3, "=INDEX($A$10:$C$12,MATCH($B$5,$A$10:$A$12,0),2)", None, 200)  # D5
    wb.close()


def _build_infer_raises_branch_limit(path: Path) -> None:
    """Build a workbook whose infer step hits the branch limit.

    A1 = OFFSET(Sheet1!B1, 0, Sheet1!C1).
    C1 is constrained with a domain that causes branch explosion in infer.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(0, 2, 0)  # C1 constrained (large interval → branch explosion)
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,0,Sheet1!C1)", None, 10)  # A1
    wb.close()


def _build_blocked_downstream_blank_leaf(path: Path) -> None:
    """Build a workbook with a downstream blank INDIRECT leaf.

    A1 -> B1 statically.
    B1 = OFFSET(F1, C1, 0), so candidate discovery must infer F1/F2 targets.
    F2 = INDIRECT(G1), and G1 is blank.

    If B1's target inference fails, a robust candidate scan should still surface G1
    rather than silently skipping the downstream dynamic-ref subgraph.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(0, 0, "=Sheet1!B1", None, 0)  # A1
    ws.write_formula(0, 1, "=OFFSET(Sheet1!F1,Sheet1!C1,0)", None, 0)  # B1
    ws.write_number(0, 2, 1)  # C1 controls which F-row is selected
    ws.write_number(0, 5, 999)  # F1 static base
    ws.write_formula(1, 5, "=INDIRECT(Sheet1!G1)", None, 0)  # F2
    wb.close()


def _make_env(mapping: dict[str, CellType]) -> CellTypeEnv:
    return mapping


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


def test_no_dynamic_refs_returns_empty(tmp_path: Path) -> None:
    """Formula with no OFFSET/INDIRECT/INDEX returns empty list."""
    path = tmp_path / "no_dyn.xlsx"
    _build_no_dynamic_refs(path)
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"])
    assert result == []


def test_single_offset_missing_leaf(tmp_path: Path) -> None:
    """Single OFFSET with one unconstrained leaf returns that leaf's address."""
    path = tmp_path / "single_offset.xlsx"
    _build_single_offset_missing_leaf(path)
    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    assert result == ["Sheet1!C1"]


def test_two_offsets_missing_leaves_collected_in_one_call(tmp_path: Path) -> None:
    """Collect missing leaves from multiple OFFSET formulas in one call.

    Two separate OFFSET formulas each with missing leaves are both collected
    in a single call — the core regression vs. the current raise-on-first behavior.
    """
    path = tmp_path / "two_offsets.xlsx"
    _build_two_offsets_missing_leaves(path)
    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    assert result == ["Sheet1!E1", "Sheet1!G1"]


def test_all_leaves_constrained_returns_empty(tmp_path: Path) -> None:
    """When all dynamic-ref leaves are constrained, returns empty list."""
    path = tmp_path / "constrained.xlsx"
    _build_all_leaves_constrained(path)
    env = _make_env(
        {
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1})),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    assert result == []


def test_partially_constrained_returns_only_missing(tmp_path: Path) -> None:
    """Only unconstrained leaves are returned; already-constrained leaves are excluded."""
    path = tmp_path / "partial.xlsx"
    _build_partial_constraint(path)
    # C1 is constrained; D1 is not
    env = _make_env(
        {
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0})),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    assert result == ["Sheet1!D1"]


def test_dynamic_refs_none_returns_all_candidates(tmp_path: Path) -> None:
    """With dynamic_refs=None all dynamic-ref leaf candidates are returned."""
    path = tmp_path / "dyn_none.xlsx"
    _build_single_offset_missing_leaf(path)
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=None)
    assert result == ["Sheet1!C1"]


def test_static_index_only_no_candidates(tmp_path: Path) -> None:
    """INDEX with only literal row/column args does not produce candidates."""
    path = tmp_path / "static_index.xlsx"
    _build_static_index_only(path)
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"])
    assert result == []


def test_index_match_range_argument_expands_all_cells(tmp_path: Path) -> None:
    """MATCH lookup range contributes every cell in A10:A12 as candidates."""
    path = tmp_path / "index_match_range_arg.xlsx"
    _build_index_match_range_arg(path)

    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!D5"], dynamic_refs=None)

    assert result == ["Sheet1!A10", "Sheet1!A11", "Sheet1!A12", "Sheet1!B5"]


def test_collect_and_continue_through_static_deps(tmp_path: Path) -> None:
    """Continue BFS after collecting missing leaves from one formula.

    One formula reports missing leaves while BFS still reaches a second
    statically-reachable dynamic-ref formula and reports its leaves too.
    """
    path = tmp_path / "collect_continue.xlsx"
    _build_two_offsets_missing_leaves(path)
    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    # Both E1 (from B1's OFFSET) and G1 (from C1's OFFSET) must be present
    assert "Sheet1!E1" in result
    assert "Sheet1!G1" in result


def test_result_is_deterministically_sorted(tmp_path: Path) -> None:
    """Output is always lexicographically sorted regardless of BFS traversal order."""
    path = tmp_path / "sorted.xlsx"
    _build_two_offsets_missing_leaves(path)
    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    assert result == sorted(result)


def test_no_offset_indirect_index_returns_empty(tmp_path: Path) -> None:
    """Workbook with no dynamic refs returns [] without raising."""
    path = tmp_path / "plain.xlsx"
    _build_no_dynamic_refs(path)
    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    assert result == []


def test_infer_raises_dynamic_ref_error_is_caught(tmp_path: Path) -> None:
    """Return an empty list when infer raises DynamicRefError.

    When all leaves are constrained but infer itself raises `DynamicRefError`
    (e.g. branch limit exceeded), the function catches it and returns `[]`
    rather than propagating.
    """
    path = tmp_path / "infer_raises.xlsx"
    _build_infer_raises_branch_limit(path)
    # C1 is constrained with a large interval; max_branches=1 forces branch explosion
    env = _make_env(
        {
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=0, max=100),
            )
        }
    )
    limits = DynamicRefLimits(max_branches=1)
    config = DynamicRefConfig(cell_type_env=env, limits=limits)
    # Must not raise — branch explosion is swallowed
    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)
    assert isinstance(result, list)


@pytest.mark.xfail(reason="Issue #254: downstream blank leaf skipped after upstream infer failure")
def test_candidate_scan_surfaces_downstream_blank_leaf_despite_blocking_infer_issue_97(
    tmp_path: Path,
) -> None:
    """Surface downstream blank leaves after upstream infer failures.

    A failed upstream dynamic-ref inference should not hide a downstream blank leaf
    that later causes graph extraction to fail.
    """
    path = tmp_path / "blocked_downstream_blank_leaf.xlsx"
    _build_blocked_downstream_blank_leaf(path)
    env = _make_env(
        {
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=0, max=100),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits(max_branches=1))

    result = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1"], dynamic_refs=config)

    assert result == ["Sheet1!G1"]


def test_missing_sheet_raises_value_error(tmp_path: Path) -> None:
    """A sheet-qualified target referencing a non-existent sheet raises ValueError."""
    path = tmp_path / "missing_sheet.xlsx"
    wb = xlsxwriter.Workbook(path)
    wb.add_worksheet("Sheet1")
    wb.close()
    with pytest.raises(ValueError, match="Sheet not found"):
        list_dynamic_ref_constraint_candidates(path, ["NoSuchSheet!A1"])


def _build_named_range_offset_workbook(path: Path) -> None:
    """Workbook with a defined-name range over a column that drives an OFFSET.

    - `OFFSET_TARGETS` -> `Sheet1!$A$1:$A$2`
    - A1 = OFFSET(B1, 0, C1)  (C1 is unconstrained leaf)
    - A2 = OFFSET(B1, 0, D1)  (D1 is unconstrained leaf)
    """
    import fastpyxl
    from fastpyxl.workbook.defined_name import DefinedName

    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"].value = "=OFFSET(Sheet1!B1,0,Sheet1!C1)"
    ws["A2"].value = "=OFFSET(Sheet1!B1,0,Sheet1!D1)"
    ws["B1"].value = 10
    ws["C1"].value = 1
    ws["D1"].value = 1
    wb.defined_names.add(DefinedName("OFFSET_TARGETS", attr_text="Sheet1!$A$1:$A$2"))
    wb.save(path)
    wb.close()


def test_candidates_accept_sheet_qualified_range_target(tmp_path: Path) -> None:
    """Sheet-qualified range targets seed the same union of leaves as expanded cells."""
    path = tmp_path / "candidates_range.xlsx"
    _build_named_range_offset_workbook(path)

    via_range = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1:A2"])
    via_cells = list_dynamic_ref_constraint_candidates(path, ["Sheet1!A1", "Sheet1!A2"])

    assert via_range == via_cells
    assert via_range == ["Sheet1!C1", "Sheet1!D1"]


def test_candidates_accept_named_range_target(tmp_path: Path) -> None:
    """A defined name pointing to a range expands to the same candidates."""
    path = tmp_path / "candidates_named_range.xlsx"
    _build_named_range_offset_workbook(path)

    via_name = list_dynamic_ref_constraint_candidates(path, ["OFFSET_TARGETS"])
    assert via_name == ["Sheet1!C1", "Sheet1!D1"]


def test_candidates_unknown_name_raises_clear_error(tmp_path: Path) -> None:
    """An unrecognized bare target token raises a ValueError mentioning the token."""
    path = tmp_path / "candidates_unknown.xlsx"
    _build_named_range_offset_workbook(path)

    with pytest.raises(ValueError) as exc:
        list_dynamic_ref_constraint_candidates(path, ["MysteryName"])
    assert "MysteryName" in str(exc.value)


def _build_shared_intermediate_offset_workbook(path: Path, n_rows: int = 20) -> None:
    """Many OFFSET formulas whose argument subgraphs share intermediate `B1`.

    Layout on 'Sheet1':
      - C1 is a leaf constant; B1 = `=Sheet1!C1+1` is a shared intermediate.
      - For each row `r` in `[2, 2+n_rows)`:
        - D<r> is a leaf column-offset constant.
        - F<r> = `=OFFSET(Sheet1!$A$1, Sheet1!$B$1, Sheet1!D<r>)`.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")

    ws.write_number(0, 0, 1)  # A1 base
    ws.write_formula(0, 1, "=Sheet1!C1+1", None, 3)  # B1 intermediate
    ws.write_number(0, 2, 2)  # C1 leaf

    for i in range(n_rows):
        row = 1 + i
        ws.write_number(row, 3, (i % 5) + 1)  # D<row+1> leaf
        ws.write_formula(
            row,
            5,
            f"=OFFSET(Sheet1!$A$1,Sheet1!$B$1,Sheet1!D{row + 1})",
            None,
            42,
        )
    wb.close()


def test_candidates_reuse_shared_cell_type_cache(tmp_path: Path) -> None:
    """list_dynamic_ref_constraint_candidates must reuse env expansion across call sites.

    Regression guard for issue #463: unlike `create_dependency_graph`, this
    function did not pass `shared_cell_type_cache=`, so the shared intermediate
    `B1` was re-inferred once per dynamic-ref formula.
    """
    from unittest.mock import patch

    from excel_grapher.grapher.builder import expand_leaf_env_to_argument_env

    n_rows = 20
    path = tmp_path / "candidates_shared_intermediate.xlsx"
    _build_shared_intermediate_offset_workbook(path, n_rows=n_rows)

    env: CellTypeEnv = {
        "Sheet1!C1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=10))
    }
    for i in range(n_rows):
        env[f"Sheet1!D{2 + i}"] = CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=1, max=5),
        )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    total_calls = 0
    passed_a_shared_cache = 0
    b1_cache_hits = 0

    def tracking_expand(*args, **kwargs):
        nonlocal total_calls, passed_a_shared_cache, b1_cache_hits
        total_calls += 1
        shared_cache = kwargs.get("shared_cell_type_cache")
        if shared_cache is not None:
            passed_a_shared_cache += 1
            if "Sheet1!B1" in shared_cache:
                b1_cache_hits += 1
        return expand_leaf_env_to_argument_env(*args, **kwargs)

    with patch(
        "excel_grapher.grapher.builder.expand_leaf_env_to_argument_env",
        side_effect=tracking_expand,
    ):
        list_dynamic_ref_constraint_candidates(path, ["Sheet1!F2"], dynamic_refs=config)

    assert total_calls > 1, "fixture should exercise multiple dynamic-ref call sites"
    assert passed_a_shared_cache == total_calls, (
        "list_dynamic_ref_constraint_candidates must pass shared_cell_type_cache= "
        f"on every expand call, got {passed_a_shared_cache}/{total_calls}"
    )
    assert b1_cache_hits == total_calls - 1, (
        f"Expected B1 to be a cache hit on {total_calls - 1} of {total_calls} calls, "
        f"but got {b1_cache_hits}. shared_cell_type_cache is not being reused."
    )


# ---------------------------------------------------------------------------
# Issue #484: memoize static-ref extraction, cache worksheets, emit progress
# ---------------------------------------------------------------------------


def _build_shared_band_offsets_workbook(path: Path, *, band: int, n_offsets: int) -> list[str]:
    """Many OFFSET formulas whose arguments share one large static SUM band.

    Layout on sheet ``S`` (matches the MCVE in issue #484):
      - ``A1`` leaf used by every ``C{i}`` formula.
      - ``C1:C{band}`` = ``=$A$1+i`` (static band).
      - For each ``j`` in ``1..n_offsets``:
        - ``D{j}`` = ``=SUM($C$1:$C${band})`` (identical formula string).
        - ``E{j}`` = ``=OFFSET($B$1,MOD(INT(D{j}),3),0)``.
      - ``B1`` leaf base for OFFSET.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_number(0, 0, 1)  # A1
    ws.write_number(0, 1, 0)  # B1
    for i in range(1, band + 1):
        ws.write_formula(i - 1, 2, f"=$A$1+{i}", None, 1 + i)  # C{i}
    targets: list[str] = []
    for j in range(1, n_offsets + 1):
        ws.write_formula(j - 1, 3, f"=SUM($C$1:$C${band})", None, 0)  # D{j}
        ws.write_formula(
            j - 1,
            4,
            f"=OFFSET($B$1,MOD(INT(D{j}),3),0)",
            None,
            0,
        )  # E{j}
        targets.append(f"S!E{j}")
    wb.close()
    return targets


def test_candidates_memoize_refs_without_dynamic_across_shared_band(
    tmp_path: Path,
) -> None:
    """Identical static formulas must not be re-parsed once per OFFSET call site.

    Regression for issue #484: the candidate helper's argument-subgraph walk
    called `_refs_without_dynamic` with no memoization, so a shared
    ``SUM($C$1:$C$band)`` tip was range-parsed once per OFFSET formula.
    """
    from unittest.mock import patch

    from excel_grapher.grapher import builder as builder_mod

    band, n_offsets = 40, 20
    path = tmp_path / "shared_band_offsets.xlsx"
    targets = _build_shared_band_offsets_workbook(path, band=band, n_offsets=n_offsets)

    real_parse_ranges = builder_mod.parse_range_refs_with_spans
    sum_formula_parses = 0

    def counting_parse_ranges(formula: str, *args, **kwargs):
        nonlocal sum_formula_parses
        if f"$C$1:$C${band}" in formula or f"C1:C{band}" in formula:
            sum_formula_parses += 1
        return real_parse_ranges(formula, *args, **kwargs)

    with patch.object(builder_mod, "parse_range_refs_with_spans", counting_parse_ranges):
        result = list_dynamic_ref_constraint_candidates(path, targets)

    # Unconstrained leaf feeding the OFFSET row argument via the SUM band.
    assert "S!A1" in result

    # Without memoization this is ~n_offsets (one SUM parse per OFFSET tip).
    # With memoization the identical SUM string is parsed once.
    assert sum_formula_parses <= 2, (
        f"Expected the shared SUM band to be range-parsed at most twice "
        f"(memoized _refs_without_dynamic), got {sum_formula_parses} parses "
        f"across {n_offsets} OFFSET formulas"
    )


def test_candidates_worksheet_cache_reduces_getitem_calls(tmp_path: Path) -> None:
    """wb[sheet] must be called at most once per unique sheet during candidate scan.

    Mirrors `create_dependency_graph`'s `_ws_f_cache` (issue #484).
    """
    from unittest.mock import patch

    import fastpyxl

    band, n_offsets = 30, 10
    path = tmp_path / "candidates_ws_cache.xlsx"
    targets = _build_shared_band_offsets_workbook(path, band=band, n_offsets=n_offsets)

    original_getitem = fastpyxl.Workbook.__getitem__
    calls: list[str] = []

    def spy_getitem(self: fastpyxl.Workbook, key: str):
        calls.append(key)
        return original_getitem(self, key)

    with patch.object(fastpyxl.Workbook, "__getitem__", spy_getitem):
        list_dynamic_ref_constraint_candidates(path, targets)

    sheet_calls = calls.count("S")
    assert sheet_calls <= 1, (
        f"Expected wb['S'] to be called at most once (worksheet cache), "
        f"but it was called {sheet_calls} times"
    )


def test_candidates_emit_bfs_and_arg_progress_traces(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Long candidate scans must emit progress traces (issue #484)."""
    from excel_grapher.grapher import builder as builder_mod
    from excel_grapher.grapher.dynamic_refs import DynamicRefTraceEvent, trace_dynamic_refs

    monkeypatch.setattr(builder_mod, "_CANDIDATES_BFS_PROGRESS_INTERVAL", 5, raising=False)
    monkeypatch.setattr(builder_mod, "_CANDIDATES_ARG_PROGRESS_INTERVAL", 20, raising=False)

    band, n_offsets = 40, 15
    path = tmp_path / "candidates_progress.xlsx"
    targets = _build_shared_band_offsets_workbook(path, band=band, n_offsets=n_offsets)

    events: list[DynamicRefTraceEvent] = []
    with trace_dynamic_refs(events.append):
        list_dynamic_ref_constraint_candidates(path, targets)

    bfs_progress = [e for e in events if e.kind == "bfs-progress"]
    arg_progress = [e for e in events if e.kind == "candidates-arg-progress"]
    assert bfs_progress, "expected bfs-progress events during candidate BFS"
    assert all(e.name == "list_dynamic_ref_constraint_candidates" for e in bfs_progress)
    assert all("nodes" in e.detail and "queue" in e.detail for e in bfs_progress)
    assert arg_progress, "expected candidates-arg-progress during argument walks"
    assert all(e.name == "list_dynamic_ref_constraint_candidates" for e in arg_progress)
    assert all("visited" in e.detail for e in arg_progress)


def test_candidates_refs_memo_returns_independent_sets(tmp_path: Path) -> None:
    """Caller mutations must not poison memoized `_refs_without_dynamic` results.

    `expand_leaf_env_to_argument_env` updates its refs set in place (`refs |= …`).
    The candidate helper memoizes `_refs_without_dynamic` across that call and
    the argument-subgraph walk, so cache hits must return a fresh set — otherwise
    expand's in-place update corrupts later lookups for the same formula.
    """
    from unittest.mock import patch

    from excel_grapher.grapher import builder as builder_mod
    from excel_grapher.grapher.dynamic_refs import expand_leaf_env_to_argument_env

    path = tmp_path / "refs_memo_isolation.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_number(0, 0, 1)  # A1
    ws.write_number(1, 0, 2)  # A2
    ws.write_number(0, 1, 0)  # B1
    # Sheet-qualified so normalize matches the raw cell string (same cache key).
    ws.write_formula(0, 3, "=S!A1+S!A2", None, 3)  # D1 tip into OFFSET
    ws.write_formula(0, 4, "=OFFSET($B$1,D1,0)", None, 0)  # E1
    wb.close()

    env = _make_env(
        {
            "S!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5)),
            "S!A2": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=5)),
            "S!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=0)),
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    poison = "S!__POISON__"
    cache_was_poisoned: list[bool] = []
    real_expand = expand_leaf_env_to_argument_env

    def wrapping_expand(
        argument_refs,
        get_cell_formula,
        get_refs_from_formula,
        *args,
        **kwargs,
    ):
        formula = "=S!A1+S!A2"
        first = get_refs_from_formula(formula, "S")
        first.add(poison)
        second = get_refs_from_formula(formula, "S")
        cache_was_poisoned.append(poison in second)
        first.discard(poison)
        return real_expand(argument_refs, get_cell_formula, get_refs_from_formula, *args, **kwargs)

    with patch.object(builder_mod, "expand_leaf_env_to_argument_env", wrapping_expand):
        list_dynamic_ref_constraint_candidates(path, ["S!E1"], dynamic_refs=config)

    assert cache_was_poisoned, "expand_leaf_env_to_argument_env was not invoked"
    assert cache_was_poisoned[0] is False, (
        "memoized _refs_without_dynamic returned a shared mutable set; "
        "in-place mutation poisoned the cache for later lookups"
    )
