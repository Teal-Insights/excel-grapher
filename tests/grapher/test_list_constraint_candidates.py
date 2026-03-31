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
    """
    A1 = B1 + C1  (static formula)
    B1 = OFFSET(Sheet1!D1, 0, Sheet1!E1)   E1 missing
    C1 = OFFSET(Sheet1!F1, 0, Sheet1!G1)   G1 missing
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
    """
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


def _build_infer_raises_branch_limit(path: Path) -> None:
    """
    A1 = OFFSET(Sheet1!B1, 0, Sheet1!C1).
    C1 is constrained with a domain that causes branch explosion in infer.
    """
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(0, 2, 0)  # C1 constrained (large interval → branch explosion)
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,0,Sheet1!C1)", None, 10)  # A1
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
    """
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


def test_collect_and_continue_through_static_deps(tmp_path: Path) -> None:
    """
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
    """
    When all leaves are constrained but infer itself raises DynamicRefError
    (e.g. branch limit exceeded), the function catches it and returns [] rather
    than propagating.
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


def test_missing_sheet_raises_value_error(tmp_path: Path) -> None:
    """A sheet-qualified target referencing a non-existent sheet raises ValueError."""
    path = tmp_path / "missing_sheet.xlsx"
    wb = xlsxwriter.Workbook(path)
    wb.add_worksheet("Sheet1")
    wb.close()
    with pytest.raises(ValueError, match="Sheet not found"):
        list_dynamic_ref_constraint_candidates(path, ["NoSuchSheet!A1"])
