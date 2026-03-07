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
    IntIntervalDomain,
)
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    DynamicRefLimits,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)


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
                interval=IntIntervalDomain(min=0, max=1),
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
                interval=IntIntervalDomain(min=0, max=10),
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
                interval=IntIntervalDomain(min=0, max=1),
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


def test_expand_leaf_env_assigns_any_when_intermediate_unsupported() -> None:
    """Intermediates that cannot be inferred (e.g. unsupported function) get CellKind.ANY.

    The cell env targets leaves; intermediates need not be constrained. When an
    intermediate's formula cannot be evaluated (e.g. uses VLOOKUP), we assign ANY
    so expansion succeeds; enumeration may later require a constraint for that cell.
    """
    from excel_grapher.core.cell_types import CellKind
    from excel_grapher.grapher.dynamic_refs import (
        DynamicRefLimits,
        expand_leaf_env_to_argument_env,
    )

    leaf_env = _make_env(
        {
            "Sheet1!A1": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=0, max=1),
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

