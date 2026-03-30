"""Regression tests for dynamic reference parsing (OFFSET/INDIRECT)."""
from __future__ import annotations

import warnings
from pathlib import Path

import pytest
import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    GreaterThanCell,
    IntervalDomain,
    IntIntervalDomain,
    NotEqualCell,
)
from excel_grapher.core.formula_ast import parse as parse_ast
from excel_grapher.grapher import dynamic_refs as dynamic_refs_mod
from excel_grapher.grapher.builder import _format_missing_leaves
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    DynamicRefLimits,
    FromWorkbook,
    constrain,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)
from excel_grapher.grapher.parser import parse_dynamic_range_refs_with_spans


def _build_offset_named_range_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(0, 3, 20)  # D1 target
    ws.write_formula(0, 2, "=1+1", None, 2)  # C1 (LANG) cached value = 2
    ws.write_formula(0, 0, "=OFFSET(B1,0,LANG)+OFFSET(B1,0,LANG)", None, 40)
    wb.define_name("LANG", "Sheet1!$C$1")
    wb.close()


def _build_offset_index_row_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    lookup = wb.add_worksheet("lookup")
    sheet = wb.add_worksheet("Sheet1")

    # Named range Country_list -> lookup!C4:C6
    wb.define_name("Country_list", "lookup!$C$4:$C$6")

    # Seed lookup range and the shifted target column
    lookup.write_number(3, 1, 111)  # B4
    lookup.write_number(3, 2, 222)  # C4

    # In B2, ROW()-ROW($B$2)+1 resolves to 1
    sheet.write_formula(
        1,
        1,
        "=OFFSET(INDEX(Country_list,ROW()-ROW($B$2)+1,1),0,-1)",
        None,
        111,
    )
    wb.close()


def _build_offset_with_arg_ref_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    start = wb.add_worksheet("START")

    ws.write_number(0, 1, 5)  # B1 base
    ws.write_number(0, 2, 99)  # C1 target
    start.write_number(9, 12, 1)  # M10 -> column offset of 1

    ws.write_formula(0, 0, "=OFFSET(B1,0,START!M10)", None, 99)
    wb.close()


def test_offset_with_cached_named_range_warns_once(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_named_range.xlsx"
    _build_offset_named_range_workbook(excel_path)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        graph = create_dependency_graph(
            excel_path, ["Sheet1!A1"], load_values=False, use_cached_dynamic_refs=True
        )

    deps = graph.dependencies("Sheet1!A1")
    # A1 = OFFSET(B1,0,LANG)+OFFSET(B1,0,LANG); LANG = Sheet1!C1. Deps include C1 (offset arg) and D1 (resolved target).
    assert deps == {"Sheet1!C1", "Sheet1!D1"}

    cache_warnings = [
        w for w in caught if "cached workbook values" in str(w.message)
    ]
    assert len(cache_warnings) == 1


def test_offset_index_row_resolves_named_range(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_index_row.xlsx"
    _build_offset_index_row_workbook(excel_path)

    graph = create_dependency_graph(
        excel_path, ["Sheet1!B2"], load_values=False, use_cached_dynamic_refs=True
    )
    deps = graph.dependencies("Sheet1!B2")
    assert deps == {"lookup!B4"}


def test_offset_index_clamps_row_when_resolved_index_exceeds_named_range() -> None:
    """INDEX row from ROW()-ROW(anchor) may exceed array height; clamp for static OFFSET resolution."""
    formula = "=OFFSET(INDEX(Country_list,ROW()-ROW($B$2)+1,1),0,-1)"
    out = parse_dynamic_range_refs_with_spans(
        formula,
        current_sheet="Sheet1",
        current_cell_a1="A10",
        named_ranges={},
        named_range_ranges={"Country_list": ("lookup", "C4", "C6")},
        value_resolver=None,
    )
    assert len(out) == 1
    start, end, _span, _arg_refs = out[0]
    assert start.sheet == "lookup" and end.sheet == "lookup"
    assert start.column == end.column == "B"
    assert start.row == end.row == 6


def test_offset_argument_references_are_dependencies(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_arg_ref.xlsx"
    _build_offset_with_arg_ref_workbook(excel_path)

    graph = create_dependency_graph(
        excel_path, ["Sheet1!A1"], load_values=False, use_cached_dynamic_refs=True
    )
    deps = graph.dependencies("Sheet1!A1")
    assert deps == {"Sheet1!C1", "START!M10"}


def _make_env(mapping: dict[str, CellType]) -> CellTypeEnv:
    return mapping


def test_dynamic_offset_with_small_integer_domain() -> None:
    # A1 = OFFSET(A1, B1, 0) with B1 in {0,1} should reach A1 and A2.
    formula = "=OFFSET(Sheet1!A1,Sheet1!B1,0)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=0, max=1),
            )
        }
    )

    targets = infer_dynamic_offset_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!A1", "Sheet1!A2"}


def test_dynamic_offset_base_index_infers_targets() -> None:
    # OFFSET(INDEX(Sheet1!A1:A3, Sheet1!B1, 1), 0, 0): INDEX returns A1, A2, or A3 when B1 in {1,2,3}.
    formula = "=OFFSET(INDEX(Sheet1!A1:Sheet1!A3,Sheet1!B1,1),0,0)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2, 3})),
            )
        }
    )
    targets = infer_dynamic_offset_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}


def test_dynamic_offset_index_row_expr_does_not_require_domain_for_row_col_ref() -> None:
    # INDEX row argument is ROW()-ROW(Sheet1!B106)+1; B106 appears only in ROW(B106).
    # We use the reference's row (106), not the cell value, so no numeric domain is required.
    formula = "=OFFSET(INDEX(Sheet1!A1:Sheet1!A5,ROW()-ROW(Sheet1!B106)+1,1),0,0)"
    env = _make_env({})
    targets = infer_dynamic_offset_targets(
        formula,
        current_sheet="Sheet1",
        cell_type_env=env,
        current_row=106,
        current_col=1,
    )
    assert targets == {"Sheet1!A1"}


def test_dynamic_offset_index_value_and_ref_only_requires_domain() -> None:
    # B1 appears in both ROW(B1) (ref_only) and as a value (B1); domain still required.
    formula = "=OFFSET(INDEX(Sheet1!A1:A3,ROW()-ROW(Sheet1!B1)+Sheet1!B1,1),0,0)"
    env = _make_env({})
    with pytest.raises(Exception) as exc_info:
        infer_dynamic_offset_targets(
            formula,
            current_sheet="Sheet1",
            cell_type_env=env,
            current_row=1,
            current_col=1,
        )
    assert "Missing CellType" in str(exc_info.value) or "must be numeric" in str(
        exc_info.value
    )


def test_dynamic_offset_requires_domains_for_all_leaf_cells() -> None:
    formula = "=OFFSET(Sheet1!A1,Sheet1!B1,0)"
    env = _make_env({})

    try:
        infer_dynamic_offset_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    except DynamicRefError as exc:
        assert "Missing CellType" in str(exc)
    else:
        raise AssertionError("Expected DynamicRefError for missing domain")


def test_dynamic_offset_uses_literal_and_enum_domains() -> None:
    # Row offset is a literal; column offset is a small enum.
    formula = "=OFFSET(Sheet1!A1,1,Sheet1!C1)"
    env = _make_env(
        {
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1})),
            )
        }
    )

    targets = infer_dynamic_offset_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    # Base is A1; row offset 1 -> row 2; col offset in {0,1} -> columns A or B.
    assert targets == {"Sheet1!A2", "Sheet1!B2"}


def test_dynamic_offset_respects_branch_limit() -> None:
    formula = "=OFFSET(Sheet1!A1,Sheet1!B1,0)"
    # Domain with 0..10 would require 11 branches; set max_branches small to force failure.
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=0, max=10),
            )
        }
    )

    limits = DynamicRefLimits(max_branches=8)
    try:
        infer_dynamic_offset_targets(
            formula,
            current_sheet="Sheet1",
            cell_type_env=env,
            limits=limits,
        )
    except DynamicRefError as exc:
        msg = str(exc)
        assert "branches" in msg or "exceeds branch limit" in msg
    else:
        raise AssertionError("Expected DynamicRefError for branch limit")


def test_dynamic_offset_argument_formulas_over_domains() -> None:
    # A1 = OFFSET(A1, SUM(B1:B3), 0) with each Bi in {0,1}.
    # SUM(B1:B3) ranges over {0,1,2,3}, so reachable rows are 1..4.
    formula = "=OFFSET(Sheet1!A1,SUM(Sheet1!B1:Sheet1!B3),0)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1})),
            ),
            "Sheet1!B2": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1})),
            ),
            "Sheet1!B3": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1})),
            ),
        }
    )

    limits = DynamicRefLimits(max_branches=64)
    targets = infer_dynamic_offset_targets(
        formula,
        current_sheet="Sheet1",
        cell_type_env=env,
        limits=limits,
    )
    assert targets == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3", "Sheet1!A4"}


def test_dynamic_offset_respects_cell_limit() -> None:
    # With B1 in {0,1,2,3}, reachable rows are {1,2,3,4}.
    formula = "=OFFSET(Sheet1!A1,Sheet1!B1,0)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1, 2, 3})),
            )
        }
    )

    limits = DynamicRefLimits(max_branches=16, max_cells=2)
    try:
        infer_dynamic_offset_targets(
            formula,
            current_sheet="Sheet1",
            cell_type_env=env,
            limits=limits,
        )
    except DynamicRefError as exc:
        msg = str(exc)
        assert "cells" in msg or "exceed limit" in msg
    else:
        raise AssertionError("Expected DynamicRefError for cell limit")


def test_dynamic_offset_respects_expr_max_depth() -> None:
    # Deeply nested IF expression in the row argument should trip max_depth.
    formula = "=OFFSET(Sheet1!A1,IF(TRUE,IF(TRUE,IF(TRUE,1,0),0),0),0)"
    env = _make_env({})

    limits = DynamicRefLimits(max_branches=1, max_cells=10, max_depth=1)
    try:
        infer_dynamic_offset_targets(
            formula,
            current_sheet="Sheet1",
            cell_type_env=env,
            limits=limits,
        )
    except DynamicRefError as exc:
        msg = str(exc)
        assert "Unsupported argument expression" in msg or "max_depth" in msg
    else:
        raise AssertionError("Expected DynamicRefError for expr depth limit")


def test_dynamic_indirect_with_literal_text() -> None:
    formula = '=INDIRECT("Sheet1!A1")'
    env = _make_env({})

    targets = infer_dynamic_indirect_targets(
        formula,
        current_sheet="Sheet1",
        cell_type_env=env,
    )
    assert targets == {"Sheet1!A1"}


def test_dynamic_indirect_over_enum_text_domain() -> None:
    # A1 = INDIRECT(B1) where B1 can point to A1 or B2.
    formula = "=INDIRECT(Sheet1!B1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.STRING,
                enum=EnumDomain(values=frozenset({"Sheet1!A1", "Sheet1!B2"})),
            )
        }
    )

    limits = DynamicRefLimits(max_branches=8)
    targets = infer_dynamic_indirect_targets(
        formula,
        current_sheet="Sheet1",
        cell_type_env=env,
        limits=limits,
    )
    assert targets == {"Sheet1!A1", "Sheet1!B2"}


def test_create_dependency_graph_raises_on_dynamic_refs_by_default(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_named_range.xlsx"
    _build_offset_named_range_workbook(excel_path)

    with pytest.raises(DynamicRefError):
        # Default behavior should not silently fall back to cached values.
        create_dependency_graph(excel_path, ["Sheet1!A1"], load_values=False)


def test_create_dependency_graph_with_dynamic_ref_config(tmp_path: Path) -> None:
    """Graph built with DynamicRefConfig resolves OFFSET and yields expected edges."""
    excel_path = tmp_path / "offset_constraint.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 0)  # B1 base
    ws.write_number(0, 2, 1)  # C1 = row offset (domain 0 or 1)
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,Sheet1!C1,0)", None, 0)  # A1
    wb.close()

    env = _make_env(
        {
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=0, max=1),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1"],
        load_values=False,
        dynamic_refs=config,
    )
    deps = graph.dependencies("Sheet1!A1")
    assert deps == {"Sheet1!B1", "Sheet1!C1", "Sheet1!B2"}


def test_create_dependency_graph_with_dynamic_ref_config_and_no_dynamic_calls(
    tmp_path: Path,
) -> None:
    """Graph building with DynamicRefConfig should work when formula has no OFFSET/INDIRECT."""
    excel_path = tmp_path / "no_dynamic_calls.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 1)  # B1
    ws.write_number(0, 2, 2)  # C1
    ws.write_formula(0, 0, "=Sheet1!B1+Sheet1!C1", None, 3)  # A1
    wb.close()

    env = _make_env({})
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1"],
        load_values=False,
        dynamic_refs=config,
    )
    deps = graph.dependencies("Sheet1!A1")
    assert deps == {"Sheet1!B1", "Sheet1!C1"}


def test_create_dependency_graph_constrain_leaf_only_formula_in_chain(tmp_path: Path) -> None:
    """Constraints are only for leaves; formula cells in the argument chain are expanded from leaf env."""
    excel_path = tmp_path / "offset_leaf_constraint.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 0)  # B1 base
    ws.write_string(9, 11, "English")  # L10 leaf (language name)
    ws.write_formula(9, 12, '=IF(L10="English",0,1)', None, 0)  # M10 formula -> 0 or 1
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,Sheet1!M10,0)", None, 0)  # A1
    wb.close()

    # Constrain only the leaf L10; M10 (formula) is expanded from this.
    env = _make_env(
        {
            "Sheet1!L10": CellType(
                kind=CellKind.STRING,
                enum=EnumDomain(values=frozenset({"English", "Other"})),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1"],
        load_values=False,
        dynamic_refs=config,
    )
    deps = graph.dependencies("Sheet1!A1")
    assert deps == {"Sheet1!B1", "Sheet1!M10", "Sheet1!B2"}


def test_create_dependency_graph_raises_for_empty_leaf_cell_missing_constraint(
    tmp_path: Path,
) -> None:
    """Empty cells that feed OFFSET are treated as leaves; missing constraint raises DynamicRefError."""
    excel_path = tmp_path / "offset_empty_leaf.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    # A1 base (has value); B1 left empty — it is the OFFSET rows argument
    ws.write_number(0, 0, 0)  # A1
    # B1 not written → empty cell
    ws.write_formula(0, 2, "=OFFSET(Sheet1!A1,Sheet1!B1,0)", None, 0)  # C1
    wb.close()

    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())

    with pytest.raises(DynamicRefError) as exc_info:
        create_dependency_graph(
            excel_path,
            ["Sheet1!C1"],
            load_values=False,
            dynamic_refs=config,
        )
    assert "Sheet1!B1" in str(exc_info.value)
    assert "leaf" in str(exc_info.value).lower()


def test_create_dependency_graph_raises_when_leaf_missing_constraint(tmp_path: Path) -> None:
    """Raises DynamicRefError listing missing leaves when only a formula cell in the chain is constrained."""
    excel_path = tmp_path / "offset_missing_leaf.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 0)  # B1 base
    ws.write_string(9, 11, "English")  # L10 leaf
    ws.write_formula(9, 12, '=IF(L10="English",0,1)', None, 0)  # M10 formula
    ws.write_formula(0, 0, "=OFFSET(Sheet1!B1,Sheet1!M10,0)", None, 0)  # A1
    wb.close()

    # Constrain M10 (formula cell) but not L10 (leaf); should raise for missing leaf.
    env = _make_env(
        {
            "Sheet1!M10": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1})),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    with pytest.raises(DynamicRefError) as exc_info:
        create_dependency_graph(
            excel_path,
            ["Sheet1!A1"],
            load_values=False,
            dynamic_refs=config,
        )
    assert "Sheet1!L10" in str(exc_info.value)
    assert "leaf" in str(exc_info.value).lower()


def test_dynamic_ref_arg_subgraph_aligns_ast_range_cap_with_builder_bfs_issue_56(
    tmp_path: Path,
) -> None:
    """Interior cells of oversized static ranges must not be required only by AST collection.

    The builder BFS uses ``expand_range(..., max_cells=...)`` (corners only when over the
    cap). Argument-env expansion must use the same cap when collecting range addresses
    from the parsed AST (GitHub issue #56).
    """
    excel_path = tmp_path / "offset_sum_range_cap_issue_56.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)  # A1
    # B1 interior of A1:C1 — left empty / unconstrained
    ws.write_number(0, 2, 1)  # C1
    ws.write_formula(0, 3, "=SUM(Sheet1!A1:C1)", None, 2)  # D1
    ws.write_number(0, 5, 0)  # F1 OFFSET base
    ws.write_formula(0, 4, "=OFFSET(Sheet1!F1,Sheet1!D1,0)", None, 0)  # E1
    wb.close()

    env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1}))
            ),
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1}))
            ),
            "Sheet1!F1": CellType(
                kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({0}))
            ),
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!E1"],
        load_values=False,
        dynamic_refs=config,
        max_range_cells=2,
    )
    assert "Sheet1!B1" not in graph


def test_dynamic_ref_missing_multiple_leaves_raises_builder_aggregate_not_per_leaf_issue_56(
    tmp_path: Path,
) -> None:
    """Several unconstrained leaves in one OFFSET argument should surface in one builder error."""
    excel_path = tmp_path / "offset_three_leaves_issue_56.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 0)  # A1 base
    ws.write_formula(
        0,
        4,
        "=OFFSET(Sheet1!A1,Sheet1!B1+Sheet1!C1+Sheet1!D1,0)",
        None,
        0,
    )  # E1
    wb.close()
    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())
    with pytest.raises(DynamicRefError) as exc_info:
        create_dependency_graph(
            excel_path,
            ["Sheet1!E1"],
            load_values=False,
            dynamic_refs=config,
        )
    msg = str(exc_info.value)
    assert "following leaf" in msg
    assert "have no constraint" in msg
    assert "Missing constraint for leaf" not in msg
    # _format_missing_leaves may merge B1:D1 into one rectangle
    assert "Sheet1!B1" in msg and "Sheet1!D1" in msg


def test_expand_leaf_env_mutual_formula_refs_in_argument_subgraph_issue_54() -> None:
    """Issue #54: mutual formula refs in the argument chain must not abort type expansion.

    Mirrors LIC-style patterns (e.g. aggregate cell referenced by scaled rows that also
    feed formulas pointing back). The cycle edge is approximated as ``ANY`` so graph build
    can continue.
    """
    leaf_env = _make_env({})
    f_a1 = "=Sheet1!B1+1"
    f_b1 = "=Sheet1!A1+1"

    def _get_cell_formula(addr: str) -> str | None:
        if addr == "Sheet1!A1":
            return f_a1
        if addr == "Sheet1!B1":
            return f_b1
        return None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {f_a1: {"Sheet1!B1"}, f_b1: {"Sheet1!A1"}}[formula]

    env = expand_leaf_env_to_argument_env(
        {"Sheet1!A1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(),
    )
    assert env["Sheet1!A1"].kind is CellKind.ANY
    assert env["Sheet1!B1"].kind is CellKind.ANY


def test_expand_leaf_env_assigns_any_when_intermediate_unsupported() -> None:
    """Intermediates that cannot be inferred (e.g. unsupported function) get CellKind.ANY.

    The cell env targets leaves; intermediates need not be constrained. When an
    intermediate's formula cannot be evaluated (e.g. uses VLOOKUP), we assign ANY
    so expansion succeeds; enumeration may later require a constraint for that cell.
    """
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=0, max=1),
            )
        }
    )

    argument_refs = {"Sheet1!B2"}

    def _get_cell_formula(addr: str) -> str | None:
        return "=FOO(Sheet1!A1)" if addr == "Sheet1!B2" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1"} if "FOO" in formula else set()

    env = expand_leaf_env_to_argument_env(
        argument_refs,
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(),
    )
    assert "Sheet1!B2" in env
    assert env["Sheet1!B2"].kind is CellKind.ANY


def test_indirect_raises_when_argument_leaf_missing_domain() -> None:
    """INDIRECT(ref) raises when ref is a leaf (non-formula) and ref has no domain in env."""
    formula = "=INDIRECT(Sheet1!B1)"
    env = _make_env({})  # B1 not in env; B1 is the leaf referenced by INDIRECT text arg

    with pytest.raises(DynamicRefError) as exc_info:
        infer_dynamic_indirect_targets(
            formula,
            current_sheet="Sheet1",
            cell_type_env=env,
        )
    msg = str(exc_info.value)
    assert "B1" in msg or "Sheet1!B1" in msg
    assert "Missing" in msg or "interval or enum" in msg


def test_indirect_does_not_raise_when_argument_is_intermediate_with_domain() -> None:
    """INDIRECT(ref) does not raise when ref is an intermediate (formula) cell with enum in env.

    The env is the expanded argument env: formula cells in the chain have domains computed
    from leaf evaluation, so they have interval or enum. Only leaves need user constraints.
    """
    formula = "=INDIRECT(Sheet1!B2)"
    # B2 would be a formula cell; expanded env gives it enum (e.g. sheet-qualified ref strings).
    env = _make_env(
        {
            "Sheet1!B2": CellType(
                kind=CellKind.STRING,
                enum=EnumDomain(values=frozenset({"Sheet1!A1", "Sheet1!B2"})),
            )
        }
    )

    targets = infer_dynamic_indirect_targets(
        formula,
        current_sheet="Sheet1",
        cell_type_env=env,
    )
    # No raise; targets are the resolved cells (same as test_dynamic_indirect_over_enum_text_domain).
    assert targets == {"Sheet1!A1", "Sheet1!B2"}


def test_format_missing_leaves_groups_contiguous_column_into_range() -> None:
    leaves = {"Sheet1!C4", "Sheet1!C5", "Sheet1!C6"}
    formatted = _format_missing_leaves(leaves)
    assert formatted == ["Sheet1!C4:Sheet1!C6"]


def test_format_missing_leaves_does_not_bridge_gaps() -> None:
    leaves = {"lookup!C4", "lookup!C73"}
    formatted = _format_missing_leaves(leaves)
    assert formatted == ["lookup!C4", "lookup!C73"]


def test_format_missing_leaves_merges_adjacent_columns_same_row() -> None:
    """Same row across consecutive columns → one horizontal range."""
    leaves = {"S!AA100", "S!AB100", "S!AC100"}
    assert _format_missing_leaves(leaves) == ["S!AA100:S!AC100"]


def test_format_missing_leaves_merges_rectangle_when_row_runs_match() -> None:
    """Identical vertical runs in adjacent columns → one rectangle per run."""
    leaves = {
        "S!AA10",
        "S!AA11",
        "S!AB10",
        "S!AB11",
    }
    assert _format_missing_leaves(leaves) == ["S!AA10:S!AB11"]


def test_format_missing_leaves_does_not_merge_columns_with_different_row_patterns() -> None:
    leaves = {"S!AA100", "S!AB101"}
    formatted = _format_missing_leaves(leaves)
    assert formatted == ["S!AA100", "S!AB101"]


def test_format_missing_leaves_splits_band_when_middle_column_differs() -> None:
    """AA and AC share rows but AB does not → no single rectangle covering AA–AC."""
    leaves = {"S!AA10", "S!AB99", "S!AC10"}
    formatted = _format_missing_leaves(leaves)
    assert formatted == ["S!AA10", "S!AB99", "S!AC10"]


def _build_simple_constant_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1
    ws.write_string(1, 1, "Afghanistan")  # B2
    wb.close()


def test_from_constraints_and_workbook_uses_workbook_values_for_constants(tmp_path: Path) -> None:
    from typing import Annotated, TypedDict

    excel_path = tmp_path / "constants.xlsx"
    _build_simple_constant_workbook(excel_path)

    class ConstConstraints(TypedDict, total=False):
        pass

    ConstConstraints.__annotations__["Sheet1!B1"] = Annotated[int, FromWorkbook()]
    ConstConstraints.__annotations__["Sheet1!B2"] = Annotated[str, FromWorkbook()]

    config = DynamicRefConfig.from_constraints_and_workbook(ConstConstraints, excel_path)
    env = config.cell_type_env

    assert env["Sheet1!B1"].kind is CellKind.NUMBER
    assert env["Sheet1!B1"].enum == EnumDomain(values=frozenset({10}))

    assert env["Sheet1!B2"].kind is CellKind.STRING
    assert env["Sheet1!B2"].enum == EnumDomain(values=frozenset({"Afghanistan"}))


def test_constrain_sets_single_cell_annotation() -> None:
    from typing import Literal, TypedDict

    class Constraints(TypedDict, total=False):
        pass

    constrain(Constraints, "Sheet1!B2", Literal["English"])

    assert Constraints.__annotations__["Sheet1!B2"] == Literal["English"]


def test_constrain_sets_all_cells_in_range() -> None:
    from typing import Literal, TypedDict

    class Constraints(TypedDict, total=False):
        pass

    constrain(Constraints, "lookup!BB4:BC5", Literal["English", "French"])

    expected_keys = {"lookup!BB4", "lookup!BC4", "lookup!BB5", "lookup!BC5"}
    assert expected_keys <= set(Constraints.__annotations__.keys())
    for key in expected_keys:
        assert Constraints.__annotations__[key] == Literal["English", "French"]


def test_constrain_accepts_quoted_sheet_range() -> None:
    from typing import Literal, TypedDict

    class Constraints(TypedDict, total=False):
        pass

    constrain(Constraints, "'Chart Data'!I21:I22", Literal[1])

    assert Constraints.__annotations__["'Chart Data'!I21"] == Literal[1]
    assert Constraints.__annotations__["'Chart Data'!I22"] == Literal[1]


def test_constrain_requires_sheet_qualified_address() -> None:
    from typing import Literal, TypedDict

    class Constraints(TypedDict, total=False):
        pass

    with pytest.raises(DynamicRefError):
        constrain(Constraints, "B2", Literal["English"])


# ── Standalone INDEX inference ──────────────────────────────────────────


def test_dynamic_index_literal_row_col() -> None:
    """INDEX with literal row and col resolves to a single cell."""
    formula = "=INDEX(Sheet1!A1:Sheet1!C3,2,3)"
    env = _make_env({})
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!C2"}


def test_dynamic_index_enum_row() -> None:
    """INDEX with constrained row enum resolves to union of cells."""
    formula = "=INDEX(Sheet1!A1:Sheet1!A3,Sheet1!B1,1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2, 3})),
            )
        }
    )
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}


def test_dynamic_index_interval_row() -> None:
    """INDEX with interval domain on row resolves to union of cells."""
    formula = "=INDEX(Sheet1!A1:Sheet1!A5,Sheet1!B1,1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                interval=IntervalDomain(min=1, max=3),
            )
        }
    )
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}


def test_dynamic_index_row_expr_with_row_function() -> None:
    """INDEX with ROW()-based expression; ref_only arg does not require domain."""
    formula = "=INDEX(Sheet1!A1:Sheet1!A5,ROW()-ROW(Sheet1!B106)+1,1)"
    env = _make_env({})
    targets = infer_dynamic_index_targets(
        formula,
        current_sheet="Sheet1",
        cell_type_env=env,
        current_row=106,
        current_col=1,
    )
    assert targets == {"Sheet1!A1"}


def test_dynamic_index_requires_domain_for_leaf() -> None:
    """INDEX with non-literal row that has no domain raises DynamicRefError."""
    formula = "=INDEX(Sheet1!A1:Sheet1!A3,Sheet1!B1,1)"
    env = _make_env({})
    with pytest.raises(DynamicRefError) as exc_info:
        infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert "Missing CellType" in str(exc_info.value) or "B1" in str(exc_info.value)


def test_dynamic_index_two_args() -> None:
    """INDEX with only 2 args (array, row_num) defaults col to 1."""
    formula = "=INDEX(Sheet1!A1:Sheet1!A3,Sheet1!B1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2})),
            )
        }
    )
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!A1", "Sheet1!A2"}


def test_dynamic_index_does_not_duplicate_with_offset_index() -> None:
    """Standalone INDEX inference ignores INDEX that is nested inside OFFSET."""
    formula = "=OFFSET(INDEX(Sheet1!A1:Sheet1!A3,Sheet1!B1,1),0,0)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2, 3})),
            )
        }
    )
    # Should return empty — INDEX inside OFFSET is not standalone
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == set()


def _build_standalone_index_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 100)  # A1
    ws.write_number(1, 0, 200)  # A2
    ws.write_number(2, 0, 300)  # A3
    ws.write_number(0, 1, 1)  # B1 (leaf: row selector)
    ws.write_formula(0, 2, "=INDEX(Sheet1!A1:Sheet1!A3,Sheet1!B1,1)", None, 100)  # C1
    wb.close()


def test_create_dependency_graph_with_standalone_index(tmp_path: Path) -> None:
    """Graph built with DynamicRefConfig resolves standalone INDEX and yields expected edges."""
    excel_path = tmp_path / "standalone_index.xlsx"
    _build_standalone_index_workbook(excel_path)

    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2, 3})),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!C1"],
        load_values=False,
        dynamic_refs=config,
    )
    deps = graph.dependencies("Sheet1!C1")
    # C1 depends on B1 (row selector) and the resolved INDEX targets A1, A2, A3
    assert deps == {"Sheet1!B1", "Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}


def test_create_dependency_graph_raises_on_standalone_index_by_default(tmp_path: Path) -> None:
    """Standalone INDEX with non-literal args raises DynamicRefError by default."""
    excel_path = tmp_path / "standalone_index_raise.xlsx"
    _build_standalone_index_workbook(excel_path)

    with pytest.raises(DynamicRefError):
        create_dependency_graph(excel_path, ["Sheet1!C1"], load_values=False)


def test_create_dependency_graph_standalone_index_missing_leaf(tmp_path: Path) -> None:
    """Standalone INDEX with missing leaf constraint raises DynamicRefError listing the leaf."""
    excel_path = tmp_path / "standalone_index_missing.xlsx"
    _build_standalone_index_workbook(excel_path)

    config = DynamicRefConfig(cell_type_env=_make_env({}), limits=DynamicRefLimits())

    with pytest.raises(DynamicRefError) as exc_info:
        create_dependency_graph(
            excel_path,
            ["Sheet1!C1"],
            load_values=False,
            dynamic_refs=config,
        )
    assert "Sheet1!B1" in str(exc_info.value)
    assert "leaf" in str(exc_info.value).lower()


def test_index_match_huge_lookup_array_only_needs_lookup_value_constraint() -> None:
    """MATCH lookup_array is not expanded for per-cell domains; INDEX stays sound via shape bounds."""
    formula = (
        "=INDEX(Sheet1!A1:Sheet1!A3,MATCH(Sheet1!B1,Sheet1!Z1:Sheet1!Z999999,0),1)"
    )
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({2})),
            )
        }
    )
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}


def test_match_domain_collection_skips_lookup_array_but_full_closure_keeps_upstream() -> None:
    ast = parse_ast("=MATCH(Sheet1!B1,OFFSET(Sheet1!Z1,0,0,5,1),0)")
    need = dynamic_refs_mod._collect_addresses_needing_domain(ast)
    assert need == {"Sheet1!B1"}
    closure = dynamic_refs_mod._collect_addresses(ast)
    assert "Sheet1!Z1" in closure


def test_expand_leaf_env_wide_interval_no_interval_branch_limit_error() -> None:
    from excel_grapher.grapher.dynamic_refs import expand_leaf_env_to_argument_env

    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=10**15),
            )
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=Sheet1!A1" if addr == "Sheet1!B1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1"} if "A1" in formula else set()

    env = expand_leaf_env_to_argument_env(
        {"Sheet1!B1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(),
    )
    assert env["Sheet1!B1"].kind is CellKind.NUMBER
    assert env["Sheet1!B1"].interval is not None
    assert env["Sheet1!B1"].interval.max == 10**15


def test_domain_from_cell_type_any_interval_is_numeric_domain() -> None:
    limits = DynamicRefLimits()
    lo, hi = -10**15, 10**15
    ct = CellType(
        kind=CellKind.ANY,
        interval=IntIntervalDomain(min=lo, max=hi),
    )
    d = dynamic_refs_mod._domain_from_cell_type(ct, limits)
    assert isinstance(d, dynamic_refs_mod._IntBounds)
    assert d.lo == lo and d.hi == hi


def test_domain_from_cell_type_any_int_enum_is_finite_ints() -> None:
    limits = DynamicRefLimits()
    ct = CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({1, 2, 3})),
    )
    d = dynamic_refs_mod._domain_from_cell_type(ct, limits)
    assert isinstance(d, dynamic_refs_mod._FiniteInts)
    assert d.values == frozenset({1, 2, 3})


def test_domain_from_cell_type_any_non_integral_float_enum_rejected() -> None:
    limits = DynamicRefLimits()
    ct = CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({1.5, 2.5})),
    )
    assert dynamic_refs_mod._domain_from_cell_type(ct, limits) is None


def test_domain_from_cell_type_any_mixed_type_enum_rejected() -> None:
    limits = DynamicRefLimits()
    ct = CellType(
        kind=CellKind.ANY,
        enum=EnumDomain(values=frozenset({1, "N/A"})),
    )
    assert dynamic_refs_mod._domain_from_cell_type(ct, limits) is None


def test_expand_leaf_env_if_isnumber_wide_any_interval_summarizes_to_number() -> None:
    """Nullable-style ANY + IntInterval on a leaf feeds IF(ISNUMBER(x),x,0) without enumeration.

    Covers the LIC-DSF-style guard (`IF(ISNUMBER(x), x, 0)`) without hitting
    `_interval_to_values` / branch-limit errors on wide bounds.
    """
    from excel_grapher.grapher.dynamic_refs import expand_leaf_env_to_argument_env

    lo, hi = -10**15, 10**15
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.ANY,
                interval=IntIntervalDomain(min=lo, max=hi),
            )
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        if addr == "Sheet1!B1":
            return "=IF(ISNUMBER(Sheet1!A1),Sheet1!A1,0)"
        return None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1"}

    env = expand_leaf_env_to_argument_env(
        {"Sheet1!B1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(),
    )
    out = env["Sheet1!B1"]
    assert out.kind is CellKind.NUMBER
    assert out.interval is not None
    assert out.interval.min == lo and out.interval.max == hi


def test_expand_leaf_env_sum_wide_range_no_branch_limit_error() -> None:
    """SUM over a range of wide-interval cells infers bounds without enumerating each interval."""
    from excel_grapher.grapher.dynamic_refs import expand_leaf_env_to_argument_env

    hi = 10**15
    financial = CellType(
        kind=CellKind.ANY,
        interval=IntIntervalDomain(min=0, max=hi),
    )
    leaf_env = _make_env(
        {
            f"Sheet1!{c}1": financial
            for c in ("A", "B", "C", "D", "E", "F")
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        if addr == "Sheet1!G1":
            return "=SUM(Sheet1!A1:Sheet1!F1)"
        return None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {f"Sheet1!{c}1" for c in ("A", "B", "C", "D", "E", "F")}

    env = expand_leaf_env_to_argument_env(
        {"Sheet1!G1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(),
    )
    out = env["Sheet1!G1"]
    assert out.kind is CellKind.NUMBER
    assert out.interval is not None
    assert out.interval.min == 0 and out.interval.max == 6 * hi


def test_expand_leaf_env_reports_unsupported_construct_in_fallback_error() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=10**9),
            )
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=ROUND(Sheet1!A1,0)" if addr == "Sheet1!B1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1"} if "A1" in formula else set()

    with pytest.raises(DynamicRefError) as exc_info:
        dynamic_refs_mod.expand_leaf_env_to_argument_env(
            {"Sheet1!B1"},
            _get_cell_formula,
            _get_refs_from_formula,
            leaf_env,
            DynamicRefLimits(max_branches=8),
        )

    msg = str(exc_info.value)
    assert "ROUND" in msg
    assert "not covered by numeric abstract analysis" in msg


def test_expand_leaf_env_division_wide_interval_no_branch_limit_error() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=10**15),
            )
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=Sheet1!A1/2" if addr == "Sheet1!B1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1"} if "A1" in formula else set()

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!B1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(max_branches=8),
    )
    out = env["Sheet1!B1"]
    assert out.kind is CellKind.NUMBER
    assert out.interval is not None
    assert out.interval.min == 0
    assert out.interval.max == 500_000_000_000_000


def test_expand_leaf_env_comparison_infers_zero_one_domain() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=-10**9, max=10**9),
            )
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=Sheet1!A1>0" if addr == "Sheet1!B1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1"} if "A1" in formula else set()

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!B1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(max_branches=8),
    )
    out = env["Sheet1!B1"]
    assert out.kind is CellKind.NUMBER
    assert out.enum is not None
    assert out.enum.values == frozenset({0, 1})


def test_expand_leaf_env_mutual_refs_terminates_with_any() -> None:
    """Mutual formula-only refs: cycle edge is ANY; expansion finishes (issue #54)."""

    def _get_cell_formula(addr: str) -> str | None:
        if addr == "Sheet1!B1":
            return "=Sheet1!C1"
        if addr == "Sheet1!C1":
            return "=Sheet1!B1"
        return None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        if "C1" in formula:
            return {"Sheet1!C1"}
        if "B1" in formula:
            return {"Sheet1!B1"}
        return set()

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!B1"},
        _get_cell_formula,
        _get_refs_from_formula,
        _make_env({}),
        DynamicRefLimits(max_depth=4),
    )
    assert env["Sheet1!B1"].kind is CellKind.ANY
    assert env["Sheet1!C1"].kind is CellKind.ANY


def test_expand_leaf_env_long_formula_chain_is_not_limited_by_expr_max_depth() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1})),
            )
        }
    )

    formulas = {
        "Sheet1!B1": "=Sheet1!A1+1",
        "Sheet1!C1": "=Sheet1!B1+1",
        "Sheet1!D1": "=Sheet1!C1+1",
        "Sheet1!E1": "=Sheet1!D1+1",
        "Sheet1!F1": "=Sheet1!E1+1",
    }

    def _get_cell_formula(addr: str) -> str | None:
        return formulas.get(addr)

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        if "A1" in formula:
            return {"Sheet1!A1"}
        if "B1" in formula:
            return {"Sheet1!B1"}
        if "C1" in formula:
            return {"Sheet1!C1"}
        if "D1" in formula:
            return {"Sheet1!D1"}
        if "E1" in formula:
            return {"Sheet1!E1"}
        return set()

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!F1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(max_depth=1),
    )

    out = env["Sheet1!F1"]
    assert out.kind is CellKind.NUMBER
    assert out.enum is not None
    assert out.enum.values == frozenset({6})


def test_expand_leaf_env_uses_ast_range_cells_when_ref_collector_only_reports_endpoints() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=10),
            ),
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=10),
            ),
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=10),
            ),
            "Sheet1!D1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=10),
            ),
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=SUM(Sheet1!A1:Sheet1!D1)" if addr == "Sheet1!E1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1", "Sheet1!D1"}

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!E1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(max_branches=8),
    )

    out = env["Sheet1!E1"]
    assert out.kind is CellKind.NUMBER
    assert out.interval is not None
    assert out.interval.min == 0
    assert out.interval.max == 40


def test_expand_leaf_env_choose_unions_selected_numeric_options() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 3})),
            ),
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({10})),
            ),
            "Sheet1!C1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Sheet1!D1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({30})),
            ),
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=CHOOSE(Sheet1!A1,Sheet1!B1,Sheet1!C1,Sheet1!D1)" if addr == "Sheet1!E1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"}

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!E1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(max_branches=8),
    )

    out = env["Sheet1!E1"]
    assert out.kind is CellKind.NUMBER
    assert out.enum is not None
    assert out.enum.values == frozenset({10, 30})


def test_infer_numeric_domain_result_reports_divisor_may_include_zero() -> None:
    limits = DynamicRefLimits(max_branches=8)
    env: CellTypeEnv = {
        "Sheet1!B9": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=0, max=50),
        ),
        "Sheet1!B10": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=1, max=100),
        ),
        "Sheet1!Q17": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=-10, max=20),
        ),
        "Sheet1!Q19": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=-10, max=20),
        ),
    }
    ast = parse_ast("=(Sheet1!Q17-Sheet1!Q19)/(Sheet1!B10-Sheet1!B9)")

    result = dynamic_refs_mod._infer_numeric_domain_result(
        ast,
        env,
        limits,
        current_sheet="Sheet1",
    )

    assert result.domain is None
    assert result.diagnostic is not None
    assert result.diagnostic.reason == "divisor may include zero"
    assert result.diagnostic.refs == frozenset({"Sheet1!B9", "Sheet1!B10"})
    assert result.diagnostic.expression == "(Sheet1!B10-Sheet1!B9)"


def test_expand_leaf_env_division_zero_risk_error_points_to_divisor_cells() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!B9": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=50),
            ),
            "Sheet1!B10": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=1, max=100),
            ),
            "Sheet1!Q17": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=-10**16, max=10**16),
            ),
            "Sheet1!Q19": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=-10**16, max=10**16),
            ),
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=(Sheet1!Q17-Sheet1!Q19)/(Sheet1!B10-Sheet1!B9)" if addr == "Sheet1!Q28" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!Q17", "Sheet1!Q19", "Sheet1!B9", "Sheet1!B10"}

    with pytest.raises(DynamicRefError) as exc_info:
        dynamic_refs_mod.expand_leaf_env_to_argument_env(
            {"Sheet1!Q28"},
            _get_cell_formula,
            _get_refs_from_formula,
            leaf_env,
            DynamicRefLimits(max_branches=8),
        )

    msg = str(exc_info.value)
    assert "divisor" in msg
    assert "include zero" in msg
    assert "Sheet1!B9" in msg
    assert "Sheet1!B10" in msg
    assert "(Sheet1!B10-Sheet1!B9)" in msg


def test_infer_numeric_domain_result_uses_relational_cell_constraint_for_divisor() -> None:
    limits = DynamicRefLimits(max_branches=8)
    env: CellTypeEnv = {
        "Sheet1!B9": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=0, max=50),
        ),
        "Sheet1!B10": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=1, max=100),
            relations=(GreaterThanCell("Sheet1!B9"),),
        ),
        "Sheet1!Q17": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=-10, max=20),
        ),
        "Sheet1!Q19": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=-10, max=20),
        ),
    }
    ast = parse_ast("=(Sheet1!Q17-Sheet1!Q19)/(Sheet1!B10-Sheet1!B9)")

    result = dynamic_refs_mod._infer_numeric_domain_result(
        ast,
        env,
        limits,
        current_sheet="Sheet1",
    )

    assert result.diagnostic is None
    assert result.domain == dynamic_refs_mod._IntBounds(-30, 30)


def test_infer_numeric_domain_result_greater_than_relation_matches_quoted_sheet_ref() -> None:
    """rel.other is normalized in metadata; parsed refs keep Excel quoting (gh #44)."""
    limits = DynamicRefLimits(max_branches=8)
    sheet = "Input 4 - External Financing"
    env: CellTypeEnv = {
        f"{sheet}!B9": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=0, max=50),
        ),
        f"{sheet}!B10": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=1, max=100),
            relations=(GreaterThanCell(f"{sheet}!B9"),),
        ),
        f"{sheet}!Q17": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=-10, max=20),
        ),
        f"{sheet}!Q19": CellType(
            kind=CellKind.NUMBER,
            interval=IntIntervalDomain(min=-10, max=20),
        ),
    }
    ast = parse_ast(f"=('{sheet}'!Q17-'{sheet}'!Q19)/('{sheet}'!B10-'{sheet}'!B9)")

    result = dynamic_refs_mod._infer_numeric_domain_result(
        ast,
        env,
        limits,
        current_sheet=sheet,
    )

    assert result.diagnostic is None
    assert result.domain == dynamic_refs_mod._IntBounds(-30, 30)


def test_infer_numeric_domain_result_uses_not_equal_constraint_for_exact_divisor() -> None:
    limits = DynamicRefLimits(max_branches=8)
    env: CellTypeEnv = {
        "Sheet1!A1": CellType(
            kind=CellKind.NUMBER,
            enum=EnumDomain(values=frozenset({1, 3})),
            relations=(NotEqualCell("Sheet1!B1"),),
        ),
        "Sheet1!B1": CellType(
            kind=CellKind.NUMBER,
            enum=EnumDomain(values=frozenset({1, 2})),
        ),
    }
    ast = parse_ast("=6/(Sheet1!A1-Sheet1!B1)")

    result = dynamic_refs_mod._infer_numeric_domain_result(
        ast,
        env,
        limits,
        current_sheet="Sheet1",
    )

    assert result.diagnostic is None
    assert result.domain == dynamic_refs_mod._FiniteInts(frozenset({-6, 3, 6}))


def test_expand_leaf_env_division_relational_constraint_avoids_zero_risk_error() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!B9": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=50),
            ),
            "Sheet1!B10": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=1, max=100),
                relations=(GreaterThanCell("Sheet1!B9"),),
            ),
            "Sheet1!Q17": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=-10**16, max=10**16),
            ),
            "Sheet1!Q19": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=-10**16, max=10**16),
            ),
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=(Sheet1!Q17-Sheet1!Q19)/(Sheet1!B10-Sheet1!B9)" if addr == "Sheet1!Q28" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!Q17", "Sheet1!Q19", "Sheet1!B9", "Sheet1!B10"}

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!Q28"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(max_branches=8),
    )

    out = env["Sheet1!Q28"]
    assert out.kind is CellKind.NUMBER
    assert out.interval == IntIntervalDomain(min=-(2 * 10**16), max=2 * 10**16)


def test_expand_leaf_env_percent_expression_stays_in_abstract_path() -> None:
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=2000),
            ),
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=1),
            ),
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        return "=Sheet1!B1+(Sheet1!A1/100)%" if addr == "Sheet1!C1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        assert sheet == "Sheet1"
        return {"Sheet1!A1", "Sheet1!B1"}

    env = dynamic_refs_mod.expand_leaf_env_to_argument_env(
        {"Sheet1!C1"},
        _get_cell_formula,
        _get_refs_from_formula,
        leaf_env,
        DynamicRefLimits(max_branches=8),
    )

    out = env["Sheet1!C1"]
    assert out.kind is CellKind.NUMBER
    assert out.enum is not None
    assert out.enum.values == frozenset({0, 1})


def test_expand_leaf_env_cartesian_product_branch_limit_error() -> None:
    """Fallback product-enumeration raises DynamicRefError when the Cartesian
    product of dependency domains exceeds max_branches, instead of hanging."""
    # 5 * 5 = 25 combinations, which exceeds max_branches=8.
    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2, 3, 4, 5})),
            ),
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({10, 20, 30, 40, 50})),
            ),
        }
    )

    def _get_cell_formula(addr: str) -> str | None:
        # ROUND is not handled by numeric abstract analysis, forcing the fallback path.
        return "=ROUND(Sheet1!A1+Sheet1!B1,0)" if addr == "Sheet1!C1" else None

    def _get_refs_from_formula(formula: str, sheet: str) -> set[str]:
        return {"Sheet1!A1", "Sheet1!B1"}

    with pytest.raises(DynamicRefError) as exc_info:
        dynamic_refs_mod.expand_leaf_env_to_argument_env(
            {"Sheet1!C1"},
            _get_cell_formula,
            _get_refs_from_formula,
            leaf_env,
            DynamicRefLimits(max_branches=8),
        )

    msg = str(exc_info.value)
    assert "Sheet1!C1" in msg  # names the formula cell
    assert "25" in msg          # actual product size
    assert "8" in msg           # the limit


def test_infer_numeric_domain_parity_never_raises() -> None:
    limits = DynamicRefLimits(max_branches=64, max_depth=12)
    env: CellTypeEnv = {
        "Sheet1!A1": CellType(
            kind=CellKind.NUMBER,
            enum=EnumDomain(values=frozenset({1, 2})),
        )
    }
    ctx = {"row": 4, "column": 2}
    tails = [
        "1",
        "1.5",
        '"x"',
        "TRUE",
        "FALSE",
        "Sheet1!A1",
        "Sheet1!A1:Sheet1!A3",
        "-Sheet1!A1",
        "Sheet1!A1+Sheet1!A1",
        "Sheet1!A1-Sheet1!A1",
        "Sheet1!A1*2",
        "Sheet1!A1/2",
        "2^3",
        "2&3",
        "1=1",
        "1<2",
        "1>0",
        "1<=1",
        "1>=0",
        "1<>2",
        "IF(TRUE,1,2)",
        "IF(FALSE,1,2)",
        "SUM(Sheet1!A1)",
        "MIN(Sheet1!A1)",
        "MAX(Sheet1!A1)",
        "ABS(-1)",
        "ROW()",
        "COLUMN()",
        "ROW(Sheet1!A5)",
        "COLUMN(Sheet1!B1)",
        'CONCAT("a","b")',
        "MATCH(Sheet1!A1,Sheet1!Z1:Sheet1!Z3,0)",
    ]
    for tail in tails:
        ast = parse_ast("=" + tail)
        out = dynamic_refs_mod._infer_numeric_domain(
            ast,
            env,
            limits,
            context=ctx,
            current_sheet="Sheet1",
        )
        assert out is None or isinstance(
            out, (dynamic_refs_mod._FiniteInts, dynamic_refs_mod._IntBounds)
        )


def test_index_sparse_row_enum_stays_precise() -> None:
    formula = "=INDEX(Sheet1!A1:Sheet1!A5,Sheet1!B1,1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 3})),
            )
        }
    )
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {"Sheet1!A1", "Sheet1!A3"}


def test_index_two_dimensional_dynamic_row_and_col() -> None:
    formula = "=INDEX(Sheet1!A1:Sheet1!C3,Sheet1!B1,Sheet1!D1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2})),
            ),
            "Sheet1!D1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({2, 3})),
            ),
        }
    )
    targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
    assert targets == {
        "Sheet1!B1",
        "Sheet1!C1",
        "Sheet1!B2",
        "Sheet1!C2",
    }


def test_index_respects_max_cells_limit() -> None:
    formula = "=INDEX(Sheet1!A1:Sheet1!A3,Sheet1!B1,1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({1, 2, 3})),
            )
        }
    )
    with pytest.raises(DynamicRefError) as exc_info:
        infer_dynamic_index_targets(
            formula,
            current_sheet="Sheet1",
            cell_type_env=env,
            limits=DynamicRefLimits(max_cells=2),
        )
    assert "cells exceed limit" in str(exc_info.value).lower()


def test_offset_per_call_max_cells_limit() -> None:
    """A single OFFSET call that fans out to more cells than max_cells must raise
    DynamicRefError immediately, not accumulate an unbounded result set."""
    # OFFSET(Sheet1!A1, 0, Sheet1!B1, 1, 1) where B1 ∈ {0,1,2,3,4} → 5 distinct
    # target cells.  With max_cells=3 the check should fire.
    formula = "=OFFSET(Sheet1!A1,0,Sheet1!B1,1,1)"
    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({0, 1, 2, 3, 4})),
            )
        }
    )
    with pytest.raises(DynamicRefError) as exc_info:
        infer_dynamic_offset_targets(
            formula,
            current_sheet="Sheet1",
            cell_type_env=env,
            limits=DynamicRefLimits(max_cells=3),
        )
    assert "cells" in str(exc_info.value).lower()
    assert "exceed limit" in str(exc_info.value).lower()
