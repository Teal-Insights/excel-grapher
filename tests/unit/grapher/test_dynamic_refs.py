"""Regression tests for dynamic reference parsing (OFFSET/INDIRECT)."""

from __future__ import annotations

import random
import warnings
from pathlib import Path
from typing import Any, Literal, cast

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
from excel_grapher.core.formula_ast import AstNode, FunctionCallNode
from excel_grapher.core.formula_ast import parse as parse_ast
from excel_grapher.grapher import dynamic_refs as dynamic_refs_mod
from excel_grapher.grapher import parser as parser_mod
from excel_grapher.grapher.builder import _format_missing_leaves
from excel_grapher.grapher.dependency_provenance import DependencyCause
from excel_grapher.grapher.dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    DynamicRefLimits,
    FromWorkbook,
    _apply_constraint_to_schema,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)
from excel_grapher.grapher.parser import FormulaNormalizer, parse_dynamic_range_refs_with_spans


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


def _build_static_offset_workbook(path: Path, *, row_expr: str, col_expr: str) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(8, 1, 5)  # B9 base
    ws.write_number(8, 2, 42)  # C9 resolved target
    ws.write_formula(8, 3, f"=OFFSET(B9,{row_expr},{col_expr})", None, 42)  # D9
    wb.close()


def _build_static_indirect_literal_workbook(path: Path, *, target_addr: str) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(8, 2, 99)  # C9
    ws.write_formula(8, 3, f'=INDIRECT("{target_addr}")', None, 99)  # D9
    wb.close()


def test_offset_with_cached_named_range_warns_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    excel_path = tmp_path / "offset_named_range.xlsx"
    _build_offset_named_range_workbook(excel_path)
    monkeypatch.setattr(parser_mod, "_WARNED_CACHED_DYNAMIC", False)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        graph = create_dependency_graph(
            excel_path, ["Sheet1!A1"], load_values=False, use_cached_dynamic_refs=True
        )

    deps = graph.get_dependencies("Sheet1!A1")
    # A1 = OFFSET(B1,0,LANG)+OFFSET(B1,0,LANG); LANG = Sheet1!C1. Deps include C1 (offset arg) and D1 (resolved target).
    assert deps == {"Sheet1!C1", "Sheet1!D1"}

    cache_warnings = [w for w in caught if "cached workbook values" in str(w.message)]
    assert len(cache_warnings) == 1
    message = str(cache_warnings[0].message)
    assert "fixed at graph-build time" in message
    assert "dynamic_refs" in message
    assert "stale" not in message.lower()


def test_offset_index_row_resolves_named_range(tmp_path: Path) -> None:
    excel_path = tmp_path / "offset_index_row.xlsx"
    _build_offset_index_row_workbook(excel_path)

    graph = create_dependency_graph(
        excel_path, ["Sheet1!B2"], load_values=False, use_cached_dynamic_refs=True
    )
    deps = graph.get_dependencies("Sheet1!B2")
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
    deps = graph.get_dependencies("Sheet1!A1")
    assert deps == {"Sheet1!C1", "START!M10"}


def test_parse_dynamic_range_refs_uses_injected_normalizer(monkeypatch: pytest.MonkeyPatch) -> None:
    class TrackingNormalizer(FormulaNormalizer):
        def __init__(self) -> None:
            super().__init__()
            self.calls: list[tuple[str, str]] = []

        def normalize(self, formula: str, current_sheet: str) -> str:
            self.calls.append((formula, current_sheet))
            return super().normalize(formula, current_sheet)

    def fail_normalize_formula(*args: object, **kwargs: object) -> str:
        raise AssertionError("legacy normalize_formula path should not be used")

    monkeypatch.setattr(parser_mod, "normalize_formula", fail_normalize_formula)
    normalizer = TrackingNormalizer()

    out = parse_dynamic_range_refs_with_spans(
        "=OFFSET(B1,0,1)",
        current_sheet="Sheet1",
        named_ranges={},
        named_range_ranges={},
        normalizer=normalizer,
        value_resolver=None,
    )

    assert len(out) == 1
    assert normalizer.calls == [("=0", "Sheet1"), ("=1", "Sheet1")]


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
    assert "Missing CellType" in str(exc_info.value) or "must be numeric" in str(exc_info.value)


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


def test_static_offset_default_path_does_not_raise(tmp_path: Path) -> None:
    excel_path = tmp_path / "static_offset_default.xlsx"
    _build_static_offset_workbook(excel_path, row_expr="0", col_expr="1")

    graph = create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)
    deps = graph.get_dependencies("Sheet1!D9")
    assert deps == {"Sheet1!C9"}


def test_static_indirect_literal_default_path_does_not_raise(tmp_path: Path) -> None:
    excel_path = tmp_path / "static_indirect_default.xlsx"
    _build_static_indirect_literal_workbook(excel_path, target_addr="Sheet1!C9")

    graph = create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)
    deps = graph.get_dependencies("Sheet1!D9")
    assert deps == {"Sheet1!C9"}


def test_randomized_static_offset_arithmetic_args_default_path(tmp_path: Path) -> None:
    rng = random.Random(1337)
    for i in range(12):
        a = rng.randint(0, 2)
        b = rng.randint(0, 2)
        row_expr = f"{a}+{b}-{a}"
        col_expr = "1+0"
        excel_path = tmp_path / f"randomized_static_offset_{i}.xlsx"
        _build_static_offset_workbook(excel_path, row_expr=row_expr, col_expr=col_expr)

        graph = create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)
        deps = graph.get_dependencies("Sheet1!D9")
        assert deps == {f"Sheet1!C{9 + b}"}


def test_randomized_static_indirect_literals_default_path(tmp_path: Path) -> None:
    rng = random.Random(7331)
    for i in range(8):
        row = rng.randint(7, 12)
        col = rng.choice(["B", "C", "D"])
        target_addr = f"Sheet1!{col}{row}"

        excel_path = tmp_path / f"randomized_static_indirect_{i}.xlsx"
        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        col_idx = ord(col) - ord("A")
        ws.write_number(row - 1, col_idx, 1)
        ws.write_formula(8, 3, f'=INDIRECT("{target_addr}")', None, 1)  # D9
        wb.close()

        graph = create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)
        deps = graph.get_dependencies("Sheet1!D9")
        assert deps == {target_addr}


def test_create_dependency_graph_raises_on_non_literal_offset_arg_by_default(
    tmp_path: Path,
) -> None:
    excel_path = tmp_path / "offset_non_literal_arg.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(8, 1, 5)  # B9
    ws.write_number(8, 2, 1)  # C9
    ws.write_formula(8, 3, "=OFFSET(B9,0,C9)", None, 42)  # D9
    wb.close()

    with pytest.raises(DynamicRefError):
        create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)


def test_create_dependency_graph_raises_on_indirect_ref_arg_by_default(tmp_path: Path) -> None:
    excel_path = tmp_path / "indirect_ref_arg.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_string(8, 1, "Sheet1!C9")  # B9
    ws.write_number(8, 2, 10)  # C9
    ws.write_formula(8, 3, "=INDIRECT(B9)", None, 10)  # D9
    wb.close()

    with pytest.raises(DynamicRefError):
        create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)


@pytest.mark.parametrize(
    "volatile_formula",
    [
        "=NOW()",
        "=TODAY()",
        "=RAND()",
        "=RANDBETWEEN(0,1)",
        "=RANDARRAY(1,1)",
        '=INFO("osversion")',
    ],
)
def test_volatile_offset_dependency_chain_warns_and_requires_cached_resolution(
    tmp_path: Path, volatile_formula: str
) -> None:
    excel_path = tmp_path / f"volatile_offset_{abs(hash(volatile_formula))}.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(8, 1, 5)  # B9 base
    ws.write_formula(8, 2, volatile_formula, None, 0)  # C9
    ws.write_formula(8, 3, "=OFFSET(B9,C9,0)", None, 5)  # D9
    wb.close()

    with (
        pytest.warns(
            UserWarning,
            match="volatile function in OFFSET/INDIRECT dependency chain",
        ),
        pytest.raises(DynamicRefError),
    ):
        create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)


@pytest.mark.parametrize(
    "volatile_formula",
    [
        "=NOW()",
        "=TODAY()",
        "=RAND()",
        "=RANDBETWEEN(0,1)",
        "=RANDARRAY(1,1)",
        '=INFO("osversion")',
    ],
)
def test_volatile_indirect_dependency_chain_warns_and_requires_cached_resolution(
    tmp_path: Path, volatile_formula: str
) -> None:
    excel_path = tmp_path / f"volatile_indirect_{abs(hash(volatile_formula))}.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(8, 1, volatile_formula, None, "Sheet1!C9")  # B9
    ws.write_number(8, 2, 10)  # C9
    ws.write_formula(8, 3, "=INDIRECT(B9)", None, 10)  # D9
    wb.close()

    with (
        pytest.warns(
            UserWarning,
            match="volatile function in OFFSET/INDIRECT dependency chain",
        ),
        pytest.raises(DynamicRefError),
    ):
        create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)


def test_cell_in_offset_chain_does_not_emit_volatile_warning(tmp_path: Path) -> None:
    excel_path = tmp_path / "cell_offset_no_volatile_warning.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(8, 1, 5)  # B9 base
    ws.write_formula(8, 2, '=CELL("row")', None, 9)  # C9
    ws.write_formula(8, 3, "=OFFSET(B9,C9,0)", None, 5)  # D9
    wb.close()

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        with pytest.raises(DynamicRefError):
            create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)
    assert not any(
        "volatile function in OFFSET/INDIRECT dependency chain" in str(w.message) for w in caught
    )


def test_cell_in_indirect_chain_does_not_emit_volatile_warning(tmp_path: Path) -> None:
    excel_path = tmp_path / "cell_indirect_no_volatile_warning.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_formula(8, 1, '=CELL("filename")', None, "Sheet1!C9")  # B9
    ws.write_number(8, 2, 10)  # C9
    ws.write_formula(8, 3, "=INDIRECT(B9)", None, 10)  # D9
    wb.close()

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        with pytest.raises(DynamicRefError):
            create_dependency_graph(excel_path, ["Sheet1!D9"], load_values=False)
    assert not any(
        "volatile function in OFFSET/INDIRECT dependency chain" in str(w.message) for w in caught
    )


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
    deps = graph.get_dependencies("Sheet1!A1")
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
    deps = graph.get_dependencies("Sheet1!A1")
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
    deps = graph.get_dependencies("Sheet1!A1")
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
    """Oversized static ranges raise instead of silently dropping interior cells.

    The builder BFS and argument-env AST collection share `expand_range`'s
    `max_cells` budget (GitHub issue #56). Exceeding it is fail-closed (#586).
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
            "Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1}))),
            "Sheet1!C1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1}))),
            "Sheet1!F1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({0}))),
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    with pytest.raises(ValueError, match="exceeding max_cells=2"):
        create_dependency_graph(
            excel_path,
            ["Sheet1!E1"],
            load_values=False,
            dynamic_refs=config,
            max_range_cells=2,
        )


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
    # _format_missing_leaves may merge B1:D1 into one single-prefix rectangle
    assert "Sheet1!B1:D1" in msg or ("Sheet1!B1" in msg and "Sheet1!D1" in msg)


def test_expand_leaf_env_mutual_formula_refs_in_argument_subgraph_issue_54() -> None:
    """Issue #54: mutual formula refs in the argument chain must not abort type expansion.

    Mirrors LIC-style patterns (e.g. aggregate cell referenced by scaled rows that also
    feed formulas pointing back). The cycle edge is approximated as `ANY` so graph build
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
    assert formatted == ["Sheet1!C4:C6"]


def test_format_missing_leaves_does_not_bridge_gaps() -> None:
    leaves = {"lookup!C4", "lookup!C73"}
    formatted = _format_missing_leaves(leaves)
    assert formatted == ["lookup!C4", "lookup!C73"]


def test_format_missing_leaves_merges_adjacent_columns_same_row() -> None:
    """Same row across consecutive columns → one horizontal range."""
    leaves = {"S!AA100", "S!AB100", "S!AC100"}
    assert _format_missing_leaves(leaves) == ["S!AA100:AC100"]


def test_format_missing_leaves_merges_rectangle_when_row_runs_match() -> None:
    """Identical vertical runs in adjacent columns → one rectangle per run."""
    leaves = {
        "S!AA10",
        "S!AA11",
        "S!AB10",
        "S!AB11",
    }
    assert _format_missing_leaves(leaves) == ["S!AA10:AB11"]


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
    from typing import Annotated

    excel_path = tmp_path / "constants.xlsx"
    _build_simple_constant_workbook(excel_path)

    schema = {
        "Sheet1!B1": Annotated[int, FromWorkbook()],
        "Sheet1!B2": Annotated[str, FromWorkbook()],
    }

    config = DynamicRefConfig.from_constraints_and_workbook(schema, excel_path)
    env = config.cell_type_env

    assert env["Sheet1!B1"].kind is CellKind.NUMBER
    assert env["Sheet1!B1"].enum == EnumDomain(values=frozenset({10}))

    assert env["Sheet1!B2"].kind is CellKind.STRING
    assert env["Sheet1!B2"].enum == EnumDomain(values=frozenset({"Afghanistan"}))


def test_apply_constraint_to_schema_sets_single_cell_annotation() -> None:
    from typing import Literal

    schema: dict[str, Any] = {}
    _apply_constraint_to_schema(schema, "Sheet1!B2", Literal["English"])

    assert schema["Sheet1!B2"] == Literal["English"]


def test_apply_constraint_to_schema_sets_all_cells_in_range() -> None:
    from typing import Literal

    schema: dict[str, Any] = {}
    _apply_constraint_to_schema(schema, "lookup!BB4:BC5", Literal["English", "French"])

    expected_keys = {"lookup!BB4", "lookup!BC4", "lookup!BB5", "lookup!BC5"}
    assert expected_keys <= set(schema.keys())
    for key in expected_keys:
        assert schema[key] == Literal["English", "French"]


def test_apply_constraint_to_schema_accepts_quoted_sheet_range() -> None:
    from typing import Literal

    schema: dict[str, Any] = {}
    _apply_constraint_to_schema(schema, "'Chart Data'!I21:I22", Literal[1])

    assert schema["'Chart Data'!I21"] == Literal[1]
    assert schema["'Chart Data'!I22"] == Literal[1]


def test_apply_constraint_to_schema_requires_sheet_qualified_address() -> None:
    from typing import Literal

    schema: dict[str, Any] = {}
    with pytest.raises(DynamicRefError):
        _apply_constraint_to_schema(schema, "B2", Literal["English"])


def test_from_constraints_rejects_legacy_typeddict_schema() -> None:
    from typing import TypedDict

    class Legacy(TypedDict, total=False):
        pass

    with pytest.raises(TypeError, match="dict\\[str, type\\]"):
        DynamicRefConfig.from_constraints(cast(Any, Legacy), {})


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


def _build_index_match_workbook(path: Path) -> None:
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


def _build_index_literal_2d_workbook(path: Path) -> None:
    """INDEX(range, literal row, literal col); array range is static (GH-156)."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    for r in range(10, 13):
        for col_idx, col_letter in enumerate(("A", "B", "C")):
            ws.write_string(r - 1, col_idx, f"{col_letter}{r}")
    ws.write_formula(0, 0, "=INDEX($A$10:$C$12,2,2)", None, 0)  # A1
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
    deps = graph.get_dependencies("Sheet1!C1")
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


def test_create_dependency_graph_index_match_includes_full_lookup_range(tmp_path: Path) -> None:
    excel_path = tmp_path / "index_match_range_graph.xlsx"
    _build_index_match_workbook(excel_path)

    env = _make_env(
        {
            "Sheet1!B5": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Sheet1!A10": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({10})),
            ),
            "Sheet1!A11": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Sheet1!A12": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({30})),
            ),
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!D5"],
        load_values=False,
        dynamic_refs=config,
    )
    deps = graph.get_dependencies("Sheet1!D5")
    assert deps == {
        "Sheet1!A10",
        "Sheet1!A11",
        "Sheet1!A12",
        "Sheet1!B5",
        "Sheet1!B11",
    }


def test_create_dependency_graph_standalone_index_records_dynamic_index_provenance(
    tmp_path: Path,
) -> None:
    """INDEX targets and row/col selectors carry `dynamic_index` (GH-531)."""
    excel_path = tmp_path / "standalone_index_provenance.xlsx"
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
        capture_dependency_provenance=True,
    )
    for dep in ("Sheet1!B1", "Sheet1!A1", "Sheet1!A2", "Sheet1!A3"):
        prov = graph.get_edge_attrs("Sheet1!C1", dep).provenance
        assert prov is not None, f"missing provenance for C1→{dep}"
        assert DependencyCause.dynamic_index in prov.causes, (
            f"expected dynamic_index on C1→{dep}, got {prov.causes}"
        )


def test_index_match_records_dynamic_index_provenance(tmp_path: Path) -> None:
    """INDEX/MATCH result and MATCH inputs carry `dynamic_index` (GH-531)."""
    excel_path = tmp_path / "index_match_provenance.xlsx"
    _build_index_match_workbook(excel_path)
    env = _make_env(
        {
            "Sheet1!B5": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Sheet1!A10": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({10})),
            ),
            "Sheet1!A11": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Sheet1!A12": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({30})),
            ),
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!D5"],
        load_values=False,
        dynamic_refs=config,
        capture_dependency_provenance=True,
    )
    for dep in ("Sheet1!A10", "Sheet1!A11", "Sheet1!A12", "Sheet1!B5", "Sheet1!B11"):
        prov = graph.get_edge_attrs("Sheet1!D5", dep).provenance
        assert prov is not None, f"missing provenance for D5→{dep}"
        assert DependencyCause.dynamic_index in prov.causes, (
            f"expected dynamic_index on D5→{dep}, got {prov.causes}"
        )


def _build_shifted_chart_data_index_match_workbook(path: Path) -> None:
    """Two row-shifted INDEX/MATCH formulas on a Chart Data sheet."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Chart Data")
    ws.write_number(4, 0, 10)  # A5
    ws.write_number(5, 0, 20)  # A6
    ws.write_number(6, 0, 30)  # A7
    ws.write_number(4, 4, 100)  # E5
    ws.write_number(5, 4, 200)  # E6
    ws.write_number(6, 4, 300)  # E7
    ws.write_number(4, 10, 10)  # K5
    ws.write_number(5, 10, 20)  # K6
    ws.write_formula(4, 20, "=INDEX($E$5:$E$7,MATCH(K5,$A$5:$A$7,0))", None, 100)  # U5
    ws.write_formula(5, 20, "=INDEX($E$5:$E$7,MATCH(K6,$A$5:$A$7,0))", None, 200)  # U6
    wb.close()


def test_shifted_chart_data_index_match_records_dynamic_index_provenance(
    tmp_path: Path,
) -> None:
    """Row-shifted Chart Data INDEX/MATCH edges carry `dynamic_index` (GH-531)."""
    excel_path = tmp_path / "chart_data_index_match_provenance.xlsx"
    _build_shifted_chart_data_index_match_workbook(excel_path)
    env = _make_env(
        {
            "Chart Data!K5": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({10})),
            ),
            "Chart Data!K6": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Chart Data!A5": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({10})),
            ),
            "Chart Data!A6": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Chart Data!A7": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({30})),
            ),
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph = create_dependency_graph(
        excel_path,
        ["'Chart Data'!U5", "'Chart Data'!U6"],
        load_values=False,
        dynamic_refs=config,
        capture_dependency_provenance=True,
    )
    expected = {
        "'Chart Data'!U5": (
            "'Chart Data'!A5",
            "'Chart Data'!A6",
            "'Chart Data'!A7",
            "'Chart Data'!K5",
            "'Chart Data'!E5",
        ),
        "'Chart Data'!U6": (
            "'Chart Data'!A5",
            "'Chart Data'!A6",
            "'Chart Data'!A7",
            "'Chart Data'!K6",
            "'Chart Data'!E6",
        ),
    }
    for formula_key, deps in expected.items():
        got = graph.get_dependencies(formula_key)
        for dep in deps:
            assert dep in got, f"{formula_key} missing dep {dep}; got {got}"
            prov = graph.get_edge_attrs(formula_key, dep).provenance
            assert prov is not None, f"missing provenance for {formula_key}→{dep}"
            assert DependencyCause.dynamic_index in prov.causes, (
                f"expected dynamic_index on {formula_key}→{dep}, got {prov.causes}"
            )


def test_create_dependency_graph_index_literal_row_col_no_array_corner_edges_issue_156(
    tmp_path: Path,
) -> None:
    """INDEX(..., literal row/col) must not register range corners from the array arg (GH-156)."""
    excel_path = tmp_path / "index_literal_2d_issue_156.xlsx"
    _build_index_literal_2d_workbook(excel_path)

    config = DynamicRefConfig.from_constraints({}, {})
    graph = create_dependency_graph(
        excel_path,
        ["Inputs!A1"],
        load_values=False,
        dynamic_refs=config,
    )
    deps = graph.get_dependencies("Inputs!A1")
    assert deps == {"Inputs!B11"}


def test_index_match_row_dep_guard_issue_156(tmp_path: Path) -> None:
    """MATCH in INDEX row position must keep the lookup value cell (GH-156 guardrail)."""
    excel_path = tmp_path / "index_match_guard_issue_156.xlsx"
    _build_index_match_workbook(excel_path)
    env = _make_env(
        {
            "Sheet1!B5": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Sheet1!A10": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({10})),
            ),
            "Sheet1!A11": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({20})),
            ),
            "Sheet1!A12": CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset({30})),
            ),
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())
    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!D5"],
        load_values=False,
        dynamic_refs=config,
    )
    assert "Sheet1!B5" in graph.get_dependencies("Sheet1!D5")


def test_index_match_huge_lookup_array_only_needs_lookup_value_constraint() -> None:
    """MATCH lookup_array is not expanded for per-cell domains; INDEX stays sound via shape bounds."""
    formula = "=INDEX(Sheet1!A1:Sheet1!A3,MATCH(Sheet1!B1,Sheet1!Z1:Sheet1!Z999999,0),1)"
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


def test_static_match_lookup_extent_resolves_index_empty_row_column_slice() -> None:
    """INDEX(range,,k) / INDEX(range,k,) are static 1-D MATCH lookup shapes (issue #497)."""
    col_slice = parse_ast("=INDEX(Trigger!AA1:AB3,,1)")
    assert isinstance(col_slice, FunctionCallNode)
    assert dynamic_refs_mod._static_match_lookup_extent(col_slice) == 3
    ordered = dynamic_refs_mod._ordered_match_lookup_cells(col_slice, current_sheet="Out")
    assert ordered == ["Trigger!AA1", "Trigger!AA2", "Trigger!AA3"]

    row_slice = parse_ast("=INDEX(Trigger!AA1:AB3,2,)")
    assert isinstance(row_slice, FunctionCallNode)
    assert dynamic_refs_mod._static_match_lookup_extent(row_slice) == 2
    ordered_row = dynamic_refs_mod._ordered_match_lookup_cells(row_slice, current_sheet="Out")
    assert ordered_row == ["Trigger!AA2", "Trigger!AB2"]


def test_columns_rows_numeric_domain_from_static_range() -> None:
    """COLUMNS/ROWS over a static range yield a singleton integer domain."""
    limits = DynamicRefLimits()
    cols = parse_ast("=COLUMNS(Trigger!AA1:AB3)")
    rows = parse_ast("=ROWS(Trigger!AA1:AB3)")
    col_dom = dynamic_refs_mod._infer_numeric_domain(cols, {}, limits, current_sheet="Out")
    row_dom = dynamic_refs_mod._infer_numeric_domain(rows, {}, limits, current_sheet="Out")
    assert isinstance(col_dom, dynamic_refs_mod._FiniteInts)
    assert col_dom.values == frozenset({2})
    assert isinstance(row_dom, dynamic_refs_mod._FiniteInts)
    assert row_dom.values == frozenset({3})


def test_columns_ref_only_skips_range_domain_collection() -> None:
    ast = parse_ast("=COLUMNS(Trigger!AA1:AB3)")
    need = dynamic_refs_mod._collect_addresses_needing_domain(ast)
    assert need == set()


def _build_vlookup_index_empty_arg_workbook(path: Path) -> None:
    """MCVE workbook for issue #497: VLOOKUP col under nested INDEX(range,,1)."""
    wb = xlsxwriter.Workbook(path)
    lookup = wb.add_worksheet("lookup")
    for row, (name, code) in enumerate([("A", 1), ("B", 2), ("C", 3)], start=1):
        lookup.write_string(row - 1, 0, name)
        lookup.write_number(row - 1, 1, code)

    trigger = wb.add_worksheet("Trigger")
    for row, name in enumerate(["A", "B", "C"], start=1):
        trigger.write_string(row - 1, 25, name)  # Z
        trigger.write_formula(
            row - 1,
            26,
            f"=VLOOKUP(Z{row}, lookup!$A$1:$B$3, 2, FALSE)",
            None,
            row,
        )  # AA
        trigger.write_number(row - 1, 27, 10 * row)  # AB

    inputs = wb.add_worksheet("Inputs")
    inputs.write_number(0, 0, 2)  # A1

    out = wb.add_worksheet("Out")
    out.write_formula(
        0,
        0,
        "=INDEX(Trigger!AA1:AA3, MATCH(Inputs!A1, Trigger!AA1:AA3, 0))",
        None,
        2,
    )
    out.write_formula(
        1,
        0,
        "=INDEX(Trigger!AA1:AB3, MATCH(Inputs!A1, Trigger!AA1:AA3, 0), 2)",
        None,
        20,
    )
    out.write_formula(
        2,
        0,
        (
            "=INDEX(Trigger!AA1:AB3, MATCH(Inputs!A1, INDEX(Trigger!AA1:AB3,,1), 0), "
            "COLUMNS(Trigger!AA1:AB3))"
        ),
        None,
        20,
    )
    wb.close()


def test_index_match_nested_index_empty_arg_over_vlookup_column(tmp_path: Path) -> None:
    """Nested INDEX(range,,1) over a VLOOKUP column must not fail domain build (#497)."""
    from typing import Annotated, Literal

    from excel_grapher.core.cell_types import Between, RealBetween

    excel_path = tmp_path / "vlookup_index_empty_arg_any.xlsx"
    _build_vlookup_index_empty_arg_workbook(excel_path)

    financial = Annotated[float, RealBetween(0, 1e6)]
    constraints: dict[str, object] = {
        "Inputs!A1": Annotated[int, Between(1, 3)],
        "lookup!A1": Literal["A"],
        "lookup!A2": Literal["B"],
        "lookup!A3": Literal["C"],
        "lookup!B1": financial,
        "lookup!B2": financial,
        "lookup!B3": financial,
        "Trigger!Z1": Literal["A"],
        "Trigger!Z2": Literal["B"],
        "Trigger!Z3": Literal["C"],
        "Trigger!AB1": Literal[10],
        "Trigger!AB2": Literal[20],
        "Trigger!AB3": Literal[30],
    }
    config = DynamicRefConfig.from_constraints(constraints, {})

    # Controls that already work without nested INDEX(range,,1).
    for target in ("Out!A1", "Out!A2"):
        graph = create_dependency_graph(
            excel_path,
            [target],
            load_values=True,
            dynamic_refs=config,
        )
        assert target in graph

    # Bug case: nested INDEX(range,,1) + COLUMNS over a VLOOKUP first column.
    graph = create_dependency_graph(
        excel_path,
        ["Out!A3"],
        load_values=True,
        dynamic_refs=config,
    )
    assert "Out!A3" in graph
    deps = graph.get_dependencies("Out!A3")
    assert {"Trigger!AB1", "Trigger!AB2", "Trigger!AB3"} & deps


def test_static_match_lookup_extent_index_zero_over_boolean_projection() -> None:
    """INDEX((range<>0),0) has the same MATCH extent as `range` (issue #506)."""
    lookup = parse_ast("=INDEX((B!N11:N14<>0),0)")
    assert isinstance(lookup, FunctionCallNode)
    assert dynamic_refs_mod._static_match_lookup_extent(lookup) == 4

    lookup_col = parse_ast("=INDEX((B!N11:N14<>0),0,1)")
    assert dynamic_refs_mod._static_match_lookup_extent(lookup_col) == 4

    # Value-preserving ordered cells stay unavailable for projected arrays.
    assert dynamic_refs_mod._ordered_match_lookup_cells(lookup, current_sheet="B") is None


def test_static_match_lookup_extent_index_zero_over_direct_range() -> None:
    """INDEX(range,0) is a static 1-D MATCH lookup shape (Excel whole-column form)."""
    col = parse_ast("=INDEX(B!N11:N14,0)")
    assert dynamic_refs_mod._static_match_lookup_extent(col) == 4
    ordered = dynamic_refs_mod._ordered_match_lookup_cells(col, current_sheet="Out")
    assert ordered == ["B!N11", "B!N12", "B!N13", "B!N14"]


def test_match_true_index_boolean_projection_numeric_domain() -> None:
    """MATCH(TRUE, INDEX((range<>0),0), 0) abstracts to IntBounds(1, N)."""
    match = parse_ast("=MATCH(TRUE,INDEX((B!N11:N14<>0),0),0)")
    dom = dynamic_refs_mod._infer_numeric_domain(match, {}, DynamicRefLimits(), current_sheet="B")
    assert isinstance(dom, dynamic_refs_mod._IntBounds)
    assert dom.lo == 1 and dom.hi == 4


def _build_match_index_first_nonzero_workbook(path: Path) -> None:
    """MCVE workbook for issue #506: first-nonzero via MATCH/INDEX boolean projection."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("B")
    for r, year in enumerate([1, 2, 3], start=11):
        ws.write_number(r - 1, 12, year)  # M
        ws.write_number(r - 1, 8, year)  # I
        ws.write_number(r - 1, 10, 0.04 - 0.01 * (r - 11))  # K
        ws.write_formula(
            r - 1,
            13,
            f'=IFNA(XLOOKUP(M{r},$I$11:$I$13,$K$11:$K$13),"")',
            None,
            0.04 - 0.01 * (r - 11),
        )  # N
    ws.write_number(13, 12, 4)  # M14
    ws.write_blank(13, 13, None)  # N14
    ws.write_formula(
        9,
        14,
        "=INDEX(N11:N14,MATCH(TRUE,INDEX((N11:N14<>0),0),0))",
        None,
        0.04,
    )  # O10
    wb.close()


def test_index_match_first_nonzero_boolean_projection_builds_graph(tmp_path: Path) -> None:
    """First-nonzero MATCH/INDEX pattern must not raise DynamicRefError (#506)."""
    from typing import Annotated, Literal

    from excel_grapher.core.cell_types import RealBetween

    excel_path = tmp_path / "match_index_first_nonzero.xlsx"
    _build_match_index_first_nonzero_workbook(excel_path)

    unit = Annotated[float, RealBetween(0, 1)]
    constraints: dict[str, object] = {
        "B!M11": Literal[1],
        "B!M12": Literal[2],
        "B!M13": Literal[3],
        "B!M14": Literal[4],
        "B!I11": Literal[1],
        "B!I12": Literal[2],
        "B!I13": Literal[3],
        "B!K11": unit,
        "B!K12": unit,
        "B!K13": unit,
        "B!N11": unit,
        "B!N12": unit,
        "B!N13": unit,
        "B!N14": Literal[None],
    }
    config = DynamicRefConfig.from_constraints(constraints, {})
    graph = create_dependency_graph(
        excel_path,
        ["B!O10"],
        load_values=True,
        dynamic_refs=config,
    )
    assert "B!O10" in graph
    deps = graph.get_dependencies("B!O10")
    assert {"B!N11", "B!N12", "B!N13", "B!N14"} & deps


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
    lo, hi = -(10**15), 10**15
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

    lo, hi = -(10**15), 10**15
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
    leaf_env = _make_env({f"Sheet1!{c}1": financial for c in ("A", "B", "C", "D", "E", "F")})

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
                interval=IntIntervalDomain(min=-(10**9), max=10**9),
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


def test_div_numeric_domains_known_nonzero_zero_endpoint_does_not_raise() -> None:
    """known_nonzero skips the zero-span guard; zero endpoints must not be used as divisors."""
    limits = DynamicRefLimits(max_branches=8)
    # Divisor domain starts at 0 (issue #487 repro).
    result = dynamic_refs_mod._div_numeric_domains(
        dynamic_refs_mod._IntBounds(1, 5),
        dynamic_refs_mod._IntBounds(0, 10),
        limits,
        known_nonzero=True,
    )
    assert result is not None
    bounds = dynamic_refs_mod._normalize_to_bounds(result)
    assert bounds.lo <= bounds.hi

    # Symmetric case: divisor domain ends at 0.
    result_hi_zero = dynamic_refs_mod._div_numeric_domains(
        dynamic_refs_mod._IntBounds(1, 5),
        dynamic_refs_mod._IntBounds(-10, 0),
        limits,
        known_nonzero=True,
    )
    assert result_hi_zero is not None


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
                interval=IntIntervalDomain(min=-(10**16), max=10**16),
            ),
            "Sheet1!Q19": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=-(10**16), max=10**16),
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
                interval=IntIntervalDomain(min=-(10**16), max=10**16),
            ),
            "Sheet1!Q19": CellType(
                kind=CellKind.NUMBER,
                interval=IntIntervalDomain(min=-(10**16), max=10**16),
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
    """Raise DynamicRefError when Cartesian product exceeds max_branches.

    Fallback product-enumeration must fail fast instead of hanging.
    """
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
    assert "25" in msg  # actual product size
    assert "8" in msg  # the limit


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
        "EXP(1)",
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
    """Raise DynamicRefError when OFFSET exceeds max_cells.

    A single OFFSET call that fans out beyond the limit must fail immediately,
    not accumulate an unbounded result set.
    """
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


def test_constraint_dynamic_ref_expansion_not_duplicated_with_provenance(
    tmp_path: Path,
) -> None:
    """Flat OFFSET formulas should share dynamic-ref expansion with provenance."""
    excel_path = tmp_path / "offset_provenance_perf.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 0)  # B1 base
    ws.write_number(0, 2, 1)  # C1 = row offset
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

    graph, call_count = _build_graph_counting_dynamic_expansion(
        excel_path,
        ["Sheet1!A1"],
        dynamic_refs=config,
    )

    assert call_count == 1, (
        f"expand_leaf_env_to_argument_env was called {call_count} times; "
        "expected 1 (shared between extraction and provenance collection)"
    )
    assert "Sheet1!B1" in graph.get_dependencies("Sheet1!A1")
    assert "Sheet1!C1" in graph.get_dependencies("Sheet1!A1")


def test_constraint_indirect_expansion_not_duplicated_with_provenance(
    tmp_path: Path,
) -> None:
    """Flat INDIRECT formulas should share dynamic-ref expansion with provenance."""
    excel_path = tmp_path / "indirect_provenance_perf.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_string(0, 1, "Sheet1!B2")  # B1 ref text
    ws.write_number(1, 1, 7)  # B2 target
    ws.write_formula(0, 0, "=INDIRECT(Sheet1!B1)", None, 7)  # A1
    wb.close()

    env = _make_env(
        {
            "Sheet1!B1": CellType(
                kind=CellKind.STRING,
                enum=EnumDomain(values=frozenset({"Sheet1!B2"})),
            )
        }
    )
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    graph, call_count = _build_graph_counting_dynamic_expansion(
        excel_path,
        ["Sheet1!A1"],
        dynamic_refs=config,
    )

    assert call_count == 1, (
        f"expand_leaf_env_to_argument_env was called {call_count} times; "
        "expected 1 for INDIRECT with provenance enabled"
    )
    assert "Sheet1!B1" in graph.get_dependencies("Sheet1!A1")
    assert "Sheet1!B2" in graph.get_dependencies("Sheet1!A1")


def test_constraint_index_dynamic_ref_expansion_not_duplicated_with_provenance(
    tmp_path: Path,
) -> None:
    """INDEX formulas should share dynamic-ref expansion with provenance (GH-531)."""
    excel_path = tmp_path / "index_provenance_perf.xlsx"
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

    graph, call_count = _build_graph_counting_dynamic_expansion(
        excel_path,
        ["Sheet1!C1"],
        dynamic_refs=config,
    )

    assert call_count == 1, (
        f"expand_leaf_env_to_argument_env was called {call_count} times; "
        "expected 1 for INDEX with provenance enabled"
    )
    for dep in ("Sheet1!B1", "Sheet1!A1", "Sheet1!A2", "Sheet1!A3"):
        prov = graph.get_edge_attrs("Sheet1!C1", dep).provenance
        assert prov is not None
        assert DependencyCause.dynamic_index in prov.causes


def test_constraint_branch_dynamic_ref_expansion_not_duplicated_with_provenance(
    tmp_path: Path,
) -> None:
    """Recursive provenance traversal should reuse cached expansion for branch formulas."""
    excel_path = tmp_path / "if_offset_provenance_perf.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 10)  # B1 base
    ws.write_number(1, 1, 20)  # B2 resolved OFFSET target
    ws.write_number(0, 2, 1)  # C1 row offset
    ws.write_boolean(0, 3, True)  # D1 IF condition
    ws.write_formula(0, 0, "=IF(Sheet1!D1,OFFSET(Sheet1!B1,Sheet1!C1,0),0)", None, 20)  # A1
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

    graph, call_count = _build_graph_counting_dynamic_expansion(
        excel_path,
        ["Sheet1!A1"],
        dynamic_refs=config,
    )

    assert call_count == 1, (
        f"expand_leaf_env_to_argument_env was called {call_count} times; "
        "expected 1 for IF branch provenance recursion"
    )
    deps = graph.get_dependencies("Sheet1!A1")
    assert "Sheet1!D1" in deps
    assert "Sheet1!B1" in deps
    assert "Sheet1!C1" in deps
    assert "Sheet1!B2" in deps


def _build_graph_counting_dynamic_expansion(
    excel_path: Path,
    targets: list[str],
    *,
    dynamic_refs: DynamicRefConfig,
):
    from unittest.mock import patch

    original_expand = expand_leaf_env_to_argument_env
    call_count = 0

    def counting_expand(*args, **kwargs):
        nonlocal call_count
        call_count += 1
        return original_expand(*args, **kwargs)

    with (
        patch(
            "excel_grapher.grapher.builder.expand_leaf_env_to_argument_env",
            side_effect=counting_expand,
        ),
        patch(
            "excel_grapher.grapher.provenance_collect.expand_leaf_env_to_argument_env",
            side_effect=counting_expand,
        ),
    ):
        graph = create_dependency_graph(
            excel_path,
            targets,
            load_values=False,
            dynamic_refs=dynamic_refs,
            capture_dependency_provenance=True,
        )

    return graph, call_count


# ---------------------------------------------------------------------------
# Benchmark: INDEX inference blowup with many row-relative formulas
# ---------------------------------------------------------------------------


def _build_wide_index_sweep_workbook(path: Path, n_rows: int = 50) -> None:
    """Create a workbook simulating the LIC-DSF Chart Data INDEX blowup.

    Layout on 'Data' sheet:
      - A1:K22 is a static data array (11 rows × 22 cols) with numeric values.
      - For each row `r` in `[2, 2+n_rows)`, column L contains
        `=INDEX(Data!$A$1:$K$22, Data!M<r>, Data!N<r>)`
        where M<r> is a row-selector leaf and N<r> is a col-selector leaf.
      - M and N columns hold small integer constants (leaves requiring constraints).

    When all M/N leaves are constrained with a wide numeric domain, inference
    is triggered for every L-column INDEX cell.  Without caching, this means
    `n_rows` separate calls to `expand_leaf_env_to_argument_env` and
    `n_rows` separate `_emit_index_targets_from_domains` invocations that
    produce the *same* target set.
    """
    import xlsxwriter

    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Data")

    # Populate the static data array A1:K22 (11 rows × 22 cols).
    for r in range(22):
        for c in range(11):
            ws.write_number(r, c, r * 11 + c + 1)

    # For each formula row, write leaf values in M and N, and an INDEX formula in L.
    for i in range(n_rows):
        row = 1 + i  # rows 2..n_rows+1 (0-indexed: 1..n_rows)
        ws.write_number(row, 12, (i % 11) + 1)  # M<row+1> = row selector (1..11)
        ws.write_number(row, 13, (i % 22) + 1)  # N<row+1> = col selector (1..22)
        # L<row+1> = INDEX($A$1:$K$22, M<row+1>, N<row+1>)
        ws.write_formula(
            row,
            11,
            f"=INDEX(Data!$A$1:$K$22,Data!M{row + 1},Data!N{row + 1})",
            None,
            42,
        )

    wb.close()


def _literal_constraint(*values: object) -> object:
    return Literal.__getitem__(values)


def _build_repeated_index_match_workbook(
    path: Path,
    *,
    formula_count: int = 25,
    country_count: int = 20,
) -> tuple[list[str], dict[str, object]]:
    """Create repeated INDEX/MATCH formulas sharing one controller/list pair."""
    countries = [f"Country {i:03d}" for i in range(1, country_count + 1)]
    wb = xlsxwriter.Workbook(path)
    inputs = wb.add_worksheet("Inputs")
    lookup = wb.add_worksheet("Lookup")
    outputs = wb.add_worksheet("Outputs")

    inputs.write("A1", countries[0])
    for row, country in enumerate(countries, start=1):
        lookup.write(row, 0, country)
        lookup.write_number(row, 1, row * 10)

    targets: list[str] = []
    last_lookup_row = country_count + 1
    for row in range(1, formula_count + 1):
        address = f"A{row}"
        outputs.write_formula(
            address,
            "=INDEX("
            f"Lookup!$B$2:$B${last_lookup_row},"
            f"MATCH(Inputs!$A$1,Lookup!$A$2:$A${last_lookup_row},0),"
            "1"
            ")",
        )
        targets.append(f"Outputs!{address}")

    wb.close()

    constraints: dict[str, object] = {"Inputs!A1": _literal_constraint(*countries)}
    for row, country in enumerate(countries, start=2):
        constraints[f"Lookup!A{row}"] = _literal_constraint(country)
    return targets, constraints


def test_repeated_identical_index_match_reuses_dynamic_ref_expansion(tmp_path: Path) -> None:
    """Repeated identical INDEX/MATCH formulas should not redo MATCH inference."""
    from unittest.mock import patch

    formula_count = 25
    country_count = 20
    excel_path = tmp_path / "repeated_index_match.xlsx"
    targets, constraints = _build_repeated_index_match_workbook(
        excel_path,
        formula_count=formula_count,
        country_count=country_count,
    )
    config = DynamicRefConfig.from_constraints(constraints, {})

    original_infer_match = dynamic_refs_mod._infer_exact_match_position_domain
    match_infer_calls = 0

    def counting_infer_match(*args: Any, **kwargs: Any):
        nonlocal match_infer_calls
        match_infer_calls += 1
        return original_infer_match(*args, **kwargs)

    with patch.object(
        dynamic_refs_mod,
        "_infer_exact_match_position_domain",
        side_effect=counting_infer_match,
    ):
        graph = create_dependency_graph(
            excel_path,
            targets,
            load_values=True,
            dynamic_refs=config,
        )

    assert len(graph.leaf_keys()) == (2 * country_count) + 1
    for target in targets:
        deps = graph.get_dependencies(target)
        assert "Inputs!A1" in deps
        assert "Lookup!A2" in deps
        assert "Lookup!B2" in deps

    assert match_infer_calls == 1


def _build_row_sensitive_index_workbook(path: Path) -> list[str]:
    """Create identical formulas whose dynamic INDEX target depends on row context."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for row in range(3):
        ws.write_number(row, 0, row + 1)
    for row in range(2):
        ws.write_formula(row, 1, "=INDEX(Sheet1!$A$1:$A$3,ROW(),1)")
    wb.close()
    return ["Sheet1!B1", "Sheet1!B2"]


def test_repeated_index_cache_keeps_row_sensitive_formulas_separate(tmp_path: Path) -> None:
    """Identical formulas containing ROW() should still use each cell's position."""
    excel_path = tmp_path / "row_sensitive_index.xlsx"
    targets = _build_row_sensitive_index_workbook(excel_path)

    graph = create_dependency_graph(
        excel_path,
        targets,
        load_values=True,
        dynamic_refs=DynamicRefConfig(cell_type_env={}, limits=DynamicRefLimits()),
    )

    assert "Sheet1!A1" in graph.get_dependencies("Sheet1!B1")
    assert "Sheet1!A2" not in graph.get_dependencies("Sheet1!B1")
    assert "Sheet1!A2" in graph.get_dependencies("Sheet1!B2")
    assert "Sheet1!A1" not in graph.get_dependencies("Sheet1!B2")


def _build_equivalent_index_match_spellings_workbook(
    path: Path,
    *,
    anchored_first: bool,
) -> tuple[list[str], dict[str, object]]:
    """Build two INDEX/MATCH formulas that normalize identically but differ in raw text."""
    countries = [f"Country {i}" for i in range(1, 5)]
    wb = xlsxwriter.Workbook(path)
    inputs = wb.add_worksheet("Inputs")
    lookup = wb.add_worksheet("Lookup")
    outputs = wb.add_worksheet("Outputs")

    inputs.write("A1", countries[0])
    for row, country in enumerate(countries, start=1):
        lookup.write(row, 0, country)
        lookup.write_number(row, 1, row * 10)

    anchored = "=INDEX(Lookup!$B$2:$B$5,MATCH(Inputs!$A$1,Lookup!$A$2:$A$5,0),1)"
    unanchored = "=INDEX(Lookup!B2:B5,MATCH(Inputs!A1,Lookup!A2:A5,0),1)"
    first, second = (anchored, unanchored) if anchored_first else (unanchored, anchored)
    outputs.write_formula("A1", first)
    outputs.write_formula("A2", second)
    wb.close()

    constraints: dict[str, object] = {"Inputs!A1": _literal_constraint(*countries)}
    for row, country in enumerate(countries, start=2):
        constraints[f"Lookup!A{row}"] = _literal_constraint(country)
    return ["Outputs!A1", "Outputs!A2"], constraints


@pytest.mark.parametrize("anchored_first", [True, False])
def test_dynamic_dep_cache_masks_current_formula_text(
    tmp_path: Path,
    anchored_first: bool,
) -> None:
    """Normalized-equivalent formulas must mask spans from their own raw text."""
    excel_path = tmp_path / f"index_match_spellings_{int(anchored_first)}.xlsx"
    targets, constraints = _build_equivalent_index_match_spellings_workbook(
        excel_path,
        anchored_first=anchored_first,
    )
    config = DynamicRefConfig.from_constraints(constraints, {})

    graph = create_dependency_graph(
        excel_path,
        targets,
        load_values=True,
        dynamic_refs=config,
    )

    expected = {
        "Inputs!A1",
        "Lookup!A2",
        "Lookup!A3",
        "Lookup!A4",
        "Lookup!A5",
        "Lookup!B2",
        "Lookup!B3",
        "Lookup!B4",
        "Lookup!B5",
    }
    for target in targets:
        assert set(graph.get_dependencies(target)) == expected


def test_wide_index_sweep_shared_env_cache(tmp_path: Path) -> None:
    """Benchmark: many INDEX formulas should share env expansion work.

    With `n_rows` INDEX formulas each referencing a different (M<r>, N<r>) leaf
    pair, every formula triggers a separate expand_leaf_env_to_argument_env call
    (since each has a unique `all_refs` set).  However, since the leaf constraints
    and intermediate cells don't overlap, we can't directly share the top-level call.

    What we *can* optimise is `_emit_index_targets_from_domains`: when every
    formula's INDEX resolves to the same `(array_range, row_dom, col_dom)`
    triple (because the leaf constraints produce the same abstract bounds), the
    target-set generation should be cached and reused.

    This test asserts:
    1. The graph is correct (all INDEX formula cells and their leaves are present).
    2. `_emit_index_targets_from_domains` is called far fewer than `n_rows` times
       after caching is in place.
    """
    n_rows = 50
    excel_path = tmp_path / "wide_index_sweep.xlsx"
    _build_wide_index_sweep_workbook(excel_path, n_rows=n_rows)

    # Build constraints: M and N leaves get interval domains matching the full array.
    env: CellTypeEnv = {}
    for i in range(n_rows):
        row = 2 + i
        env[f"Data!M{row}"] = CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=1, max=11),
        )
        env[f"Data!N{row}"] = CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=1, max=22),
        )

    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())

    targets = [f"Data!L{2 + i}" for i in range(n_rows)]

    # Count calls to _emit_index_targets_from_domains.
    from unittest.mock import patch

    original_emit = dynamic_refs_mod._emit_index_targets_from_domains
    emit_call_count = 0

    def counting_emit(*args, **kwargs):
        nonlocal emit_call_count
        emit_call_count += 1
        return original_emit(*args, **kwargs)

    with patch.object(
        dynamic_refs_mod,
        "_emit_index_targets_from_domains",
        side_effect=counting_emit,
    ):
        graph = create_dependency_graph(
            excel_path,
            targets,
            load_values=False,
            dynamic_refs=config,
        )

    # Correctness: every target should be in the graph and have M/N leaf deps.
    for i in range(n_rows):
        row = 2 + i
        node_key = f"Data!L{row}"
        deps = graph.get_dependencies(node_key)
        assert f"Data!M{row}" in deps, f"Missing M leaf dep for {node_key}"
        assert f"Data!N{row}" in deps, f"Missing N leaf dep for {node_key}"

    # Optimization: with target-set caching, _emit_index_targets_from_domains
    # should be called only once for the unique (array_range, row_dom, col_dom).
    # Without caching it would be called n_rows times.
    assert emit_call_count < n_rows, (
        f"_emit_index_targets_from_domains called {emit_call_count} times for "
        f"{n_rows} INDEX formulas; expected caching to reduce this to ~1"
    )


def _build_shared_intermediate_index_workbook(path: Path, n_rows: int = 30) -> None:
    """Workbook where INDEX formulas share intermediate argument cells.

    Layout on 'Sheet1':
      - A1:E10 is the data array (10 rows × 5 cols).
      - B1 is a shared intermediate: `=C1+1` (formula cell, not a leaf).
      - C1 is a leaf (numeric constant = 2).
      - For each row `r` in `[2, 2+n_rows)`:
        - F<r> = `=INDEX($A$1:$E$10, Sheet1!B1, Sheet1!D<r>)`
        - D<r> is a leaf (col selector constant).

    All `n_rows` INDEX formulas share the intermediate cell B1 in their
    argument subgraph.  With a shared env expansion cache, B1's type is
    inferred once and reused across all formula cells.
    """
    import xlsxwriter

    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")

    # Data array A1:E10.
    for r in range(10):
        for c in range(5):
            ws.write_number(r, c, r * 5 + c + 1)

    # B1 = =C1+1 (intermediate formula).
    ws.write_formula(0, 1, "=Sheet1!C1+1", None, 3)
    # C1 = 2 (leaf).
    ws.write_number(0, 2, 2)

    for i in range(n_rows):
        row = 1 + i  # 0-indexed rows 1..n_rows
        ws.write_number(row, 3, (i % 5) + 1)  # D<row+1> col selector
        ws.write_formula(
            row,
            5,
            f"=INDEX(Sheet1!$A$1:$E$10,Sheet1!$B$1,Sheet1!D{row + 1})",
            None,
            42,
        )

    wb.close()


def test_shared_intermediate_env_expansion_cache(tmp_path: Path) -> None:
    """Many INDEX formulas sharing an intermediate cell should reuse env expansion.

    All formulas reference B1 (intermediate) which depends on C1 (leaf).
    With a shared env expansion cache, `cell_type_for(B1)` is computed once
    and reused across all 30 formula cells' expand_leaf_env_to_argument_env calls.

    We verify this by wrapping expand_leaf_env_to_argument_env and inspecting
    the shared_cell_type_cache: B1 should already be present in the cache for
    all calls after the first.
    """
    from unittest.mock import patch

    n_rows = 30
    excel_path = tmp_path / "shared_intermediate.xlsx"
    _build_shared_intermediate_index_workbook(excel_path, n_rows=n_rows)

    env: CellTypeEnv = {
        "Sheet1!C1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=10))
    }
    for i in range(n_rows):
        row = 2 + i
        env[f"Sheet1!D{row}"] = CellType(
            kind=CellKind.NUMBER,
            interval=IntervalDomain(min=1, max=5),
        )

    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())
    targets = [f"Sheet1!F{2 + i}" for i in range(n_rows)]

    # Track how many times B1 is already in the shared cache when
    # expand_leaf_env_to_argument_env is called.
    original_expand = expand_leaf_env_to_argument_env
    total_calls = 0

    def tracking_expand(*args, **kwargs):
        nonlocal total_calls
        total_calls += 1
        return original_expand(*args, **kwargs)

    with patch(
        "excel_grapher.grapher.builder.expand_leaf_env_to_argument_env",
        side_effect=tracking_expand,
    ):
        graph = create_dependency_graph(
            excel_path,
            targets,
            load_values=False,
            dynamic_refs=config,
        )

    # Correctness check: each formula should depend on B1 and its D leaf.
    for i in range(n_rows):
        row = 2 + i
        deps = graph.get_dependencies(f"Sheet1!F{row}")
        assert "Sheet1!B1" in deps, f"Missing B1 dep for F{row}"
        assert f"Sheet1!D{row}" in deps, f"Missing D leaf dep for F{row}"

    # Optimization check: shape-keyed INDEX inference (issue #716) shares the
    # target set, so expand runs once; B1 is inferred on that call and the
    # remaining rows reuse `_dyn_shape_cache` without re-expanding.
    assert total_calls == 1, (
        f"Expected 1 expand call for identical-shape INDEX rows, got {total_calls}"
    )


def _build_shifted_index_same_match_workbook(
    path: Path,
    *,
    formula_count: int = 12,
    country_count: int = 8,
) -> tuple[list[str], dict[str, object]]:
    """INDEX arrays shift per row; MATCH arguments stay absolute and identical.

    This is the LIC-DSF chart-series shape from issue #528: each output cell is
    a variation of `INDEX(shifted_range, MATCH(fixed_lookup, fixed_list, 0))`.
    Normalized formulae differ (so `_dyn_dep_cache` misses) while the MATCH
    argument subgraph — the only refs collected for env expansion — is the same.
    """
    countries = [f"Country {i:03d}" for i in range(1, country_count + 1)]
    wb = xlsxwriter.Workbook(path)
    inputs = wb.add_worksheet("Inputs")
    lookup = wb.add_worksheet("Lookup")
    chart = wb.add_worksheet("Chart Data")
    outputs = wb.add_worksheet("Outputs")

    inputs.write("A1", countries[0])
    last_lookup_row = country_count + 1
    for row, country in enumerate(countries, start=1):
        lookup.write(row, 0, country)
        lookup.write_number(row, 1, row * 10)

    for row in range(1, formula_count + country_count + 2):
        chart.write_number(row - 1, 0, row)
        chart.write_number(row - 1, 1, row * 100)

    targets: list[str] = []
    for i in range(formula_count):
        start = i + 1
        end = start + country_count - 1
        outputs.write_formula(
            i,
            0,
            "=INDEX("
            f"'Chart Data'!A{start}:B{end},"
            f"MATCH(Inputs!$A$1,Lookup!$A$2:$A${last_lookup_row},0),"
            "1"
            ")",
        )
        targets.append(f"Outputs!A{i + 1}")

    wb.close()
    constraints: dict[str, object] = {"Inputs!A1": _literal_constraint(*countries)}
    for row, country in enumerate(countries, start=2):
        constraints[f"Lookup!A{row}"] = _literal_constraint(country)
    return targets, constraints


def test_shifted_index_variants_reuse_identical_argument_env(tmp_path: Path) -> None:
    """Row-wise INDEX variants must not re-derive the same MATCH argument env.

    Regression for issue #528: `_dyn_cache` / `_dyn_dep_cache` key on the
    normalized formula (plus position), so a relative INDEX array makes every
    row a miss even when `expand_leaf_env_to_argument_env` would return the
    same env.  One expansion should serve every variant; INDEX target inference
    still runs per formula so shifted arrays keep distinct deps.
    """
    formula_count = 12
    country_count = 8
    excel_path = tmp_path / "shifted_index_same_match.xlsx"
    targets, constraints = _build_shifted_index_same_match_workbook(
        excel_path,
        formula_count=formula_count,
        country_count=country_count,
    )
    config = DynamicRefConfig.from_constraints(constraints, {})

    graph, call_count = _build_graph_counting_dynamic_expansion(
        excel_path,
        targets,
        dynamic_refs=config,
    )

    assert call_count == 1, (
        f"expand_leaf_env_to_argument_env was called {call_count} times for "
        f"{formula_count} INDEX variants with an identical MATCH subgraph; "
        "expected 1 (issue #528)"
    )

    first_deps = set(graph.get_dependencies(targets[0]))
    last_deps = set(graph.get_dependencies(targets[-1]))
    assert "Inputs!A1" in first_deps
    assert "Lookup!A2" in first_deps
    assert "'Chart Data'!A1" in first_deps
    assert f"'Chart Data'!A{formula_count}" in last_deps
    assert "'Chart Data'!A1" not in last_deps
    assert f"'Chart Data'!A{formula_count}" not in first_deps


def _build_shifted_offset_same_match_workbook(
    path: Path,
    *,
    formula_count: int = 12,
    country_count: int = 8,
) -> tuple[list[str], dict[str, object]]:
    """OFFSET bases shift per row; MATCH arguments stay absolute and identical.

    Sibling of `_build_shifted_index_same_match_workbook`: each output cell is
    a variation of `OFFSET(shifted_base, MATCH(fixed_lookup, fixed_list, 0)-1, 0)`.
    The OFFSET base is not an argument-subgraph ref (static A1, no nested call),
    so env expansion should still see one MATCH subgraph while OFFSET targets
    remain per-formula.
    """
    countries = [f"Country {i:03d}" for i in range(1, country_count + 1)]
    wb = xlsxwriter.Workbook(path)
    inputs = wb.add_worksheet("Inputs")
    lookup = wb.add_worksheet("Lookup")
    chart = wb.add_worksheet("Chart Data")
    outputs = wb.add_worksheet("Outputs")

    inputs.write("A1", countries[0])
    last_lookup_row = country_count + 1
    for row, country in enumerate(countries, start=1):
        lookup.write(row, 0, country)
        lookup.write_number(row, 1, row * 10)

    for row in range(1, formula_count + country_count + 2):
        chart.write_number(row - 1, 0, row)
        chart.write_number(row - 1, 1, row * 100)

    targets: list[str] = []
    for i in range(formula_count):
        start = i + 1
        outputs.write_formula(
            i,
            0,
            "=OFFSET("
            f"'Chart Data'!A{start},"
            f"MATCH(Inputs!$A$1,Lookup!$A$2:$A${last_lookup_row},0)-1,"
            "0"
            ")",
        )
        targets.append(f"Outputs!A{i + 1}")

    wb.close()
    constraints: dict[str, object] = {"Inputs!A1": _literal_constraint(*countries)}
    for row, country in enumerate(countries, start=2):
        constraints[f"Lookup!A{row}"] = _literal_constraint(country)
    return targets, constraints


def test_shifted_offset_variants_reuse_identical_argument_env(tmp_path: Path) -> None:
    """Row-wise OFFSET variants must not re-derive the same MATCH argument env.

    Provenance collection *does* analyse OFFSET (unlike INDEX), so this also
    guards against a `_dyn_cache` miss re-expanding the MATCH subgraph during
    `capture_dependency_provenance=True` (the default for this helper).
    """
    formula_count = 12
    country_count = 8
    excel_path = tmp_path / "shifted_offset_same_match.xlsx"
    targets, constraints = _build_shifted_offset_same_match_workbook(
        excel_path,
        formula_count=formula_count,
        country_count=country_count,
    )
    config = DynamicRefConfig.from_constraints(constraints, {})

    graph, call_count = _build_graph_counting_dynamic_expansion(
        excel_path,
        targets,
        dynamic_refs=config,
    )

    assert call_count == 1, (
        f"expand_leaf_env_to_argument_env was called {call_count} times for "
        f"{formula_count} OFFSET variants with an identical MATCH subgraph; "
        "expected 1 (issue #528)"
    )

    first_deps = set(graph.get_dependencies(targets[0]))
    last_deps = set(graph.get_dependencies(targets[-1]))
    assert "Inputs!A1" in first_deps
    assert "Lookup!A2" in first_deps
    assert "'Chart Data'!A1" in first_deps
    assert f"'Chart Data'!A{formula_count}" in last_deps
    assert "'Chart Data'!A1" not in last_deps
    assert f"'Chart Data'!A{formula_count}" not in first_deps


def test_argument_subgraph_memo_returns_independent_sets(tmp_path: Path) -> None:
    """Caller mutations of argument-subgraph sets must not poison the memo.

    `_argument_subgraph_refs` is cached by `frozenset(argument_addrs)`.  If the
    cache stores the live `all_refs` set, `expand_leaf_env_to_argument_env`'s
    caller (or a wrapper) mutating that set would change later lookups for the
    same MATCH subgraph and defeat env reuse.
    """
    from unittest.mock import patch

    from excel_grapher.grapher import builder as builder_mod

    formula_count = 12
    excel_path = tmp_path / "subgraph_memo_isolation.xlsx"
    targets, constraints = _build_shifted_offset_same_match_workbook(
        excel_path,
        formula_count=formula_count,
    )
    config = DynamicRefConfig.from_constraints(constraints, {})

    poison = "Outputs!ZZ999"
    original_expand = expand_leaf_env_to_argument_env
    call_count = 0

    def wrapping_expand(argument_refs, *args, **kwargs):
        nonlocal call_count
        call_count += 1
        argument_refs.add(poison)
        clean = {addr for addr in argument_refs if addr != poison}
        return original_expand(clean, *args, **kwargs)

    with (
        patch.object(builder_mod, "expand_leaf_env_to_argument_env", wrapping_expand),
        patch(
            "excel_grapher.grapher.provenance_collect.expand_leaf_env_to_argument_env",
            wrapping_expand,
        ),
    ):
        graph = create_dependency_graph(
            excel_path,
            targets,
            load_values=False,
            dynamic_refs=config,
            capture_dependency_provenance=True,
        )

    assert call_count == 1, (
        f"expand_leaf_env_to_argument_env was called {call_count} times; "
        "mutating the argument-subgraph set must not poison the memo "
        f"(expected 1 expand for {formula_count} OFFSET variants)"
    )
    assert "'Chart Data'!A1" in graph.get_dependencies(targets[0])


# ---------------------------------------------------------------------------
# Deterministic traversal order
# ---------------------------------------------------------------------------


class TestDeterministicTraversalOrder:
    """Visit cells in deterministic sorted order during env expansion.

    `expand_leaf_env_to_argument_env` must visit cells in sorted order so
    partial-progress persistence and error messages stay reproducible regardless
    of Python's randomized set iteration.

    The test instruments the function to record the order in which
    `cell_type_for` is first called for each address, then asserts that
    order is lexicographic across 50 repeated invocations.
    """

    # Diamond graph:
    #   D1 = B1 + C1   (top)
    #   B1 = A1 + 1    (left arm)
    #   C1 = A1 * 2    (right arm)
    #   A1 = leaf {1, 2, 3}

    _FORMULAS: dict[str, str] = {
        "Sheet1!D1": "=Sheet1!B1+Sheet1!C1",
        "Sheet1!B1": "=Sheet1!A1+1",
        "Sheet1!C1": "=Sheet1!A1*2",
    }
    _REFS: dict[str, set[str]] = {
        "=Sheet1!B1+Sheet1!C1": {"Sheet1!B1", "Sheet1!C1"},
        "=Sheet1!A1+1": {"Sheet1!A1"},
        "=Sheet1!A1*2": {"Sheet1!A1"},
    }
    _LEAF_ENV: CellTypeEnv = {
        "Sheet1!A1": CellType(
            kind=CellKind.NUMBER,
            enum=EnumDomain(values=frozenset({1, 2, 3})),
        ),
    }

    def _run_once(self) -> tuple[dict[str, CellType], list[str]]:
        visit_order: list[str] = []
        # Patch get_cell_formula to record first-visit order
        real_get_formula = self._FORMULAS.get

        def tracking_get_formula(addr: str) -> str | None:
            if addr not in visit_order:
                visit_order.append(addr)
            return real_get_formula(addr)

        env = expand_leaf_env_to_argument_env(
            {"Sheet1!D1"},
            tracking_get_formula,
            lambda formula, sheet: self._REFS.get(formula, set()),
            self._LEAF_ENV,
            DynamicRefLimits(),
        )
        return env, visit_order

    def test_traversal_order_is_sorted(self) -> None:
        """Visit formula cells in sorted order at each traversal level.

        This ensures deterministic behaviour across Python processes with
        different hash seeds, which is essential for reproducible error
        messages and reliable partial-progress persistence.
        """
        _env, order = self._run_once()
        # D1 is the root (requested), then its deps B1 and C1 should be
        # visited in sorted order, then A1 as their shared dependency.
        assert order[0] == "Sheet1!D1", f"Root should be visited first, got {order[0]}"
        # B1 should come before C1 (sorted)
        b1_idx = order.index("Sheet1!B1")
        c1_idx = order.index("Sheet1!C1")
        assert b1_idx < c1_idx, (
            f"Sheet1!B1 (index {b1_idx}) should be visited before "
            f"Sheet1!C1 (index {c1_idx}) in sorted traversal"
        )

    def test_inferred_domains_are_correct(self) -> None:
        """Verify the expected numeric domains for the diamond graph."""
        env, _ = self._run_once()
        # A1 = {1, 2, 3}  (leaf)
        assert env["Sheet1!A1"].enum is not None
        assert env["Sheet1!A1"].enum.values == frozenset({1, 2, 3})
        # B1 = A1 + 1 = {2, 3, 4}
        assert env["Sheet1!B1"].enum is not None
        assert env["Sheet1!B1"].enum.values == frozenset({2, 3, 4})
        # C1 = A1 * 2 = {2, 4, 6}
        assert env["Sheet1!C1"].enum is not None
        assert env["Sheet1!C1"].enum.values == frozenset({2, 4, 6})
        # D1 = B1 + C1 = {2+2, 2+4, 2+6, 3+2, 3+4, 3+6, 4+2, 4+4, 4+6}
        #              = {4, 6, 8, 5, 7, 9, 6, 8, 10} = {4, 5, 6, 7, 8, 9, 10}
        assert env["Sheet1!D1"].enum is not None
        assert env["Sheet1!D1"].enum.values == frozenset({4, 5, 6, 7, 8, 9, 10})


# ---------------------------------------------------------------------------
# Abstract-path characterization (Phase 0)
# ---------------------------------------------------------------------------


def _parse_selector(expr: str) -> AstNode:
    """Parse a bare selector expression (without the leading '=') into an AST node.

    Wraps the expression with '=' so that parse_ast can handle it, then unwraps
    the outer node that parse_ast wraps around the expression body.
    """
    return parse_ast("=" + expr)


class TestAbstractPathCharacterization:
    """Guardrail tests proving which formulas stay on the abstract path.

    These tests characterize current behaviour so that later phases can measure
    actual improvement against a known baseline.
    """

    def _limits(self) -> DynamicRefLimits:
        return DynamicRefLimits()

    # ------------------------------------------------------------------
    # Test 1: INDEX <- IF <- comparison stays abstract
    # ------------------------------------------------------------------

    def test_if_comparison_selector_stays_abstract(self) -> None:
        """IF(A1>B1, 1, 2) as row selector stays on the abstract path.

        The result is {1, 2} regardless of A1/B1 values.  This must NOT require
        Cartesian enumeration of all A1 x B1 combinations.
        """
        ast = _parse_selector("IF(Sheet1!A1>Sheet1!B1, 1, 2)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=10)),
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=10)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None, (
            "IF(A1>B1, 1, 2) should stay on the abstract path (domain not None)"
        )
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        assert result.domain.values == frozenset({1, 2})

    # ------------------------------------------------------------------
    # Test 2: constant-condition IF yields exact domain
    # ------------------------------------------------------------------

    def test_if_constant_true_condition_yields_then_domain(self) -> None:
        """IF(1, 3, 7) should return domain {3}."""
        ast = _parse_selector("IF(1, 3, 7)")
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, {}, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        assert result.domain.values == frozenset({3})

    def test_if_constant_false_condition_yields_else_domain(self) -> None:
        """IF(0, 3, 7) should return domain {7}."""
        ast = _parse_selector("IF(0, 3, 7)")
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, {}, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        assert result.domain.values == frozenset({7})

    # ------------------------------------------------------------------
    # Test 2b: Non-one non-zero conditions are provably truthy
    # ------------------------------------------------------------------

    def test_if_nonzero_integer_is_provably_true(self) -> None:
        """IF(2, 3, 7): condition domain {2} is provably truthy → result {3}."""
        ast = _parse_selector("IF(2, 3, 7)")
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, {}, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        assert result.domain.values == frozenset({3})

    def test_if_all_nonzero_domain_is_provably_true(self) -> None:
        """IF(A1, 3, 7): A1 domain {2, 4} (all non-zero) → provably truthy → result {3}."""
        ast = _parse_selector("IF(Sheet1!A1, 3, 7)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2, 4})))}
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        assert result.domain.values == frozenset({3})

    # ------------------------------------------------------------------
    # Test 3: ABS / MIN / MAX currently fall back (abstract gap baseline)
    # ------------------------------------------------------------------

    def test_abs_stays_abstract_after_phase1(self) -> None:
        """ABS(A1) stays on the abstract path after Phase 1 (regression guard)."""
        ast = _parse_selector("ABS(Sheet1!A1)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5))}
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None, "ABS should now produce an abstract domain"

    def test_exp_stays_abstract_after_phase1(self) -> None:
        """EXP(A1) stays on the abstract path after Phase 1 (regression guard)."""
        ast = _parse_selector("EXP(Sheet1!A1)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5))}
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None, "EXP should now produce an abstract domain"
        bounds = dynamic_refs_mod._normalize_to_bounds(result.domain)
        assert bounds.lo == 2
        assert bounds.hi == 149

    def test_min_stays_abstract_after_phase1(self) -> None:
        """MIN(A1, B1) stays on the abstract path after Phase 1 (regression guard)."""
        ast = _parse_selector("MIN(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5)),
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=2, max=6)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None, "MIN should now produce an abstract domain"

    def test_max_stays_abstract_after_phase1(self) -> None:
        """MAX(A1, B1) stays on the abstract path after Phase 1 (regression guard)."""
        ast = _parse_selector("MAX(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5)),
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=2, max=6)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None, "MAX should now produce an abstract domain"


# ---------------------------------------------------------------------------
# Phase 1: ABS / MIN / MAX abstract transfer rules
# ---------------------------------------------------------------------------


class TestAbstractMinMaxAbsTransferRules:
    """ABS, MIN, and MAX should stay on the abstract path after Phase 1."""

    def _limits(self) -> DynamicRefLimits:
        return DynamicRefLimits()

    # ------------------------------------------------------------------
    # ABS
    # ------------------------------------------------------------------

    def test_abs_finite_positive_returns_exact_set(self) -> None:
        """ABS with a small positive enum domain returns exact values."""
        ast = _parse_selector("ABS(Sheet1!A1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 3, 5}))
                )
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        assert result.domain.values == frozenset({1, 3, 5})

    def test_abs_finite_negative_returns_exact_set(self) -> None:
        """ABS with negative enum domain maps to absolute values."""
        ast = _parse_selector("ABS(Sheet1!A1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({-3, -1, 2}))
                )
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        assert result.domain.values == frozenset({1, 2, 3})

    def test_abs_positive_bounds_returns_same_interval(self) -> None:
        """ABS([100, 1200]) = [100, 1200] (entirely positive, span > max_branches -> _IntBounds)."""
        ast = _parse_selector("ABS(Sheet1!A1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER, interval=IntervalDomain(min=100, max=1200)
                )
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        b = dynamic_refs_mod._normalize_to_bounds(result.domain)
        assert b.lo == 100
        assert b.hi == 1200

    def test_abs_negative_bounds_returns_negated_interval(self) -> None:
        """ABS([-1200, -100]) = [100, 1200] (entirely negative, span > max_branches -> _IntBounds)."""
        ast = _parse_selector("ABS(Sheet1!A1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER, interval=IntervalDomain(min=-1200, max=-100)
                )
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        b = dynamic_refs_mod._normalize_to_bounds(result.domain)
        assert b.lo == 100
        assert b.hi == 1200

    def test_abs_mixed_bounds_returns_hull(self) -> None:
        """ABS([-600, 700]) = [0, 700] (crosses zero; hull is [0, max(600,700)])."""
        ast = _parse_selector("ABS(Sheet1!A1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER, interval=IntervalDomain(min=-600, max=700)
                )
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        b = dynamic_refs_mod._normalize_to_bounds(result.domain)
        assert b.lo == 0
        assert b.hi == 700

    def test_exp_finite_enum_returns_conservative_bounds(self) -> None:
        """EXP with a small enum domain maps to a conservative integer hull."""
        ast = _parse_selector("EXP(Sheet1!A1)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({0, 1})))}
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        b = dynamic_refs_mod._normalize_to_bounds(result.domain)
        assert b.lo == 1
        assert b.hi == 3

    def test_abs_unsupported_returns_none(self) -> None:
        """ABS with no domain info (e.g. unknown cell) still returns None."""
        ast = _parse_selector("ABS(Sheet1!A1)")
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, {}, self._limits())
        assert result.domain is None

    # ------------------------------------------------------------------
    # MIN
    # ------------------------------------------------------------------

    def test_min_two_finite_domains_returns_exact_set(self) -> None:
        """MIN(A1, B1) with small finite domains returns exact minimum set."""
        ast = _parse_selector("MIN(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2, 4}))
                ),
                "Sheet1!B1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({3, 5}))
                ),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        # min(2,3)=2, min(2,5)=2, min(4,3)=3, min(4,5)=4
        assert result.domain.values == frozenset({2, 3, 4})

    def test_min_two_bounds_returns_interval(self) -> None:
        """MIN([1,5], [2,6]) = [1, 5] (safe interval hull for bounds)."""
        ast = _parse_selector("MIN(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5)),
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=2, max=6)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        b = dynamic_refs_mod._normalize_to_bounds(result.domain)
        assert b.lo == 1
        assert b.hi == 5

    def test_min_unsupported_arg_returns_none(self) -> None:
        """MIN with an unknown/unsupported arg returns None (not a crash)."""
        ast = _parse_selector("MIN(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2})))}
            # B1 not in env -> returns None domain
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is None

    # ------------------------------------------------------------------
    # MAX
    # ------------------------------------------------------------------

    def test_max_two_finite_domains_returns_exact_set(self) -> None:
        """MAX(A1, B1) with small finite domains returns exact maximum set."""
        ast = _parse_selector("MAX(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2, 4}))
                ),
                "Sheet1!B1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({3, 5}))
                ),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        assert isinstance(result.domain, dynamic_refs_mod._FiniteInts)
        # max(2,3)=3, max(2,5)=5, max(4,3)=4, max(4,5)=5
        assert result.domain.values == frozenset({3, 4, 5})

    def test_max_two_bounds_returns_interval(self) -> None:
        """MAX([1,5], [2,6]) = [2, 6] (safe interval hull for bounds)."""
        ast = _parse_selector("MAX(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5)),
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=2, max=6)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        b = dynamic_refs_mod._normalize_to_bounds(result.domain)
        assert b.lo == 2
        assert b.hi == 6

    def test_max_unsupported_arg_returns_none(self) -> None:
        """MAX with an unknown/unsupported arg returns None (not a crash)."""
        ast = _parse_selector("MAX(Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 2})))}
            # B1 not in env -> None domain
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is None

    # ------------------------------------------------------------------
    # Integration: INDEX + ABS/MIN/MAX stays abstract (no enumeration fallback)
    # ------------------------------------------------------------------

    def test_index_abs_row_stays_abstract(self) -> None:
        """INDEX(A1:A5, ABS(B1), 1) with B1 in [1,3] should stay on the abstract path."""
        formula = "=INDEX(Sheet1!A1:Sheet1!A5, ABS(Sheet1!B1), 1)"
        env = _make_env(
            {"Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=3))}
        )
        targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
        # Abstract path: ABS([1,3]) = [1,3]; INDEX picks rows 1-3 -> A1, A2, A3
        assert targets == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}

    def test_index_min_row_stays_abstract(self) -> None:
        """INDEX(A1:A5, MIN(B1, C1), 1) with B1,C1 in [1,3] should stay on abstract path."""
        formula = "=INDEX(Sheet1!A1:Sheet1!A5, MIN(Sheet1!B1, Sheet1!C1), 1)"
        env = _make_env(
            {
                "Sheet1!B1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2, 3}))
                ),
                "Sheet1!C1": CellType(
                    kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1, 4}))
                ),
            }
        )
        targets = infer_dynamic_index_targets(formula, current_sheet="Sheet1", cell_type_env=env)
        # min(2,1)=1, min(2,4)=2, min(3,1)=1, min(3,4)=3 -> rows {1,2,3}
        assert targets == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}


# ---------------------------------------------------------------------------
# Phase 2: Branch-local environment refinement for IF
# ---------------------------------------------------------------------------


class TestBranchLocalIfRefinement:
    """Narrow branch-local domains for supported IF predicates.

    IF analysis should refine domains when the condition is in the supported
    predicate fragment (cell op literal, AND of such comparisons).
    """

    def _limits(self) -> DynamicRefLimits:
        return DynamicRefLimits()

    def test_if_diff_narrows_then_branch(self) -> None:
        """IF(A1>B1, A1-B1, 0): then-branch should only include positive differences.

        Without refinement the union would include 0 (from the else branch which yields 0,
        but also potentially 0 from A1-B1 which could be 0 if A1==B1).
        With refinement, when A1>B1 is known in the then-branch, A1-B1 > 0.
        """
        ast = _parse_selector("IF(Sheet1!A1>Sheet1!B1, Sheet1!A1-Sheet1!B1, 0)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5)),
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        domain_values = (
            result.domain.values
            if isinstance(result.domain, dynamic_refs_mod._FiniteInts)
            else set(range(result.domain.lo, result.domain.hi + 1))
        )
        # After refinement, A1>B1 in the then branch means A1-B1 >= 1 (not 0).
        # The else branch is 0.  The union should exclude negative differences.
        assert 0 in domain_values  # else branch contributes 0
        assert not any(v < 0 for v in domain_values), (
            "Branch refinement should prevent negative values in the then branch"
        )

    def test_if_narrows_index_selector_domain(self) -> None:
        """INDEX(A1:A5, IF(A1>B1, A1, B1), 1): refined branches produce a narrower domain.

        A1 in [3,5], B1 in [1,3].  Without refinement both full domains are unioned.
        With refinement:
          then-branch (A1>B1): A1 in [4,5] (must be > B1 max=3), so row in [4,5]
          else-branch (A1<=B1): B1 in [3,3] at narrowest, but A1 can be [3,3] too -> B1 up to 3
        The result should be a subset of [1..5] (correct) and the test verifies it's not empty.
        """
        ast = _parse_selector("IF(Sheet1!A1>Sheet1!B1, Sheet1!A1, Sheet1!B1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=3, max=5)),
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=3)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None

    def test_if_and_narrows_then_to_exact_range(self) -> None:
        """IF(AND(A1>=2, A1<=4), A1, 0): then-branch narrows A1 domain to {2,3,4}."""
        ast = _parse_selector("IF(AND(Sheet1!A1>=2, Sheet1!A1<=4), Sheet1!A1, 0)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=10))}
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        domain_values = (
            result.domain.values
            if isinstance(result.domain, dynamic_refs_mod._FiniteInts)
            else set(range(result.domain.lo, result.domain.hi + 1))
        )
        # then-branch: A1 in {2,3,4}; else-branch: 0  -> union {0,2,3,4}
        assert 2 in domain_values
        assert 3 in domain_values
        assert 4 in domain_values
        assert 0 in domain_values
        assert not any(v > 4 for v in domain_values), (
            "Values > 4 should be excluded by AND(A1>=2, A1<=4) refinement"
        )
        assert not any(v == 1 for v in domain_values), (
            "Value 1 should be excluded by AND(A1>=2, ...) refinement"
        )

    def test_if_then_branch_diagnostic_does_not_propagate(self) -> None:
        """Divisor-zero diagnostic from the then-branch should not propagate through IF.

        When the condition is ambiguous and the then-branch has a potential
        division-by-zero, the result should be domain=None (fall back to
        enumeration) rather than carrying the diagnostic out of the IF.
        """
        ast = _parse_selector("IF(Sheet1!X1, Sheet1!A1/Sheet1!B1, 5)")
        env = _make_env(
            {
                "Sheet1!X1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=1)),
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5)),
                # B1 includes zero → A1/B1 produces a divisor-may-include-zero diagnostic
                "Sheet1!B1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=3)),
            }
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.diagnostic is None, (
            "Divisor-zero diagnostic from then-branch must not propagate out of IF"
        )
        assert result.domain is None, "When then-branch is unsound, result domain should be None"

    def test_if_unsupported_condition_degrades_to_union(self) -> None:
        """IF(ISNUMBER(A1), 1, 2): condition not a plain comparison; falls back to union."""
        # ISNUMBER is handled but it returns {0,1} not a refineable constraint
        ast = _parse_selector("IF(ISNUMBER(Sheet1!A1), 3, 7)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({1})))}
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        # Should not crash; should return a valid union domain
        assert result.domain is not None

    def test_if_literal_gt_cell_refines_else_branch(self) -> None:
        """IF(3>A1, A1, 0): the condition is '3 > A1' i.e. A1 < 3.

        In the then branch (3>A1 is true): A1 < 3, so A1 in {1,2} from [1,5].
        In the else branch (3>A1 is false): A1 >= 3, so A1 in {3,4,5}.
        But the else branch yields 0 (a literal), so the result is union of {1,2} ∪ {0}.
        """
        ast = _parse_selector("IF(3>Sheet1!A1, Sheet1!A1, 0)")
        env = _make_env(
            {"Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=5))}
        )
        result = dynamic_refs_mod._infer_numeric_domain_result(ast, env, self._limits())
        assert result.domain is not None
        domain_values = (
            result.domain.values
            if isinstance(result.domain, dynamic_refs_mod._FiniteInts)
            else set(range(result.domain.lo, result.domain.hi + 1))
        )
        # then: A1 in {1,2}; else: 0 -> {0,1,2}
        assert domain_values == frozenset({0, 1, 2})


# ---------------------------------------------------------------------------
# Phase 3: Branch-aware domain collection
# ---------------------------------------------------------------------------


class TestBranchAwareDomainCollection:
    """_collect_addresses_needing_domain should skip refs from provably dead branches."""

    def _limits(self) -> DynamicRefLimits:
        return DynamicRefLimits()

    def test_collect_skips_dead_branch_with_constant_condition(self) -> None:
        """IF(1, A1+B1, C1+D1): condition is always true; only {A1,B1} needed."""
        ast = _parse_selector("IF(1, Sheet1!A1+Sheet1!B1, Sheet1!C1+Sheet1!D1)")
        addrs = dynamic_refs_mod._collect_addresses_needing_domain(
            ast, env={}, limits=self._limits()
        )
        assert "Sheet1!A1" in addrs
        assert "Sheet1!B1" in addrs
        assert "Sheet1!C1" not in addrs
        assert "Sheet1!D1" not in addrs

    def test_collect_includes_both_branches_for_ambiguous_condition(self) -> None:
        """IF(Sheet1!X1, A1+B1, C1+D1): condition is unknown; collect all refs."""
        ast = _parse_selector("IF(Sheet1!X1, Sheet1!A1+Sheet1!B1, Sheet1!C1+Sheet1!D1)")
        env = _make_env(
            {"Sheet1!X1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=0, max=1))}
        )
        addrs = dynamic_refs_mod._collect_addresses_needing_domain(
            ast, env=env, limits=self._limits()
        )
        assert "Sheet1!A1" in addrs
        assert "Sheet1!B1" in addrs
        assert "Sheet1!C1" in addrs
        assert "Sheet1!D1" in addrs

    def test_collect_choose_by_constant_index_prunes_dead_branches(self) -> None:
        """CHOOSE(2, A1+B1, C1+D1): index is 2; only {C1,D1} needed."""
        ast = _parse_selector("CHOOSE(2, Sheet1!A1+Sheet1!B1, Sheet1!C1+Sheet1!D1)")
        addrs = dynamic_refs_mod._collect_addresses_needing_domain(
            ast, env={}, limits=self._limits()
        )
        assert "Sheet1!C1" in addrs
        assert "Sheet1!D1" in addrs
        assert "Sheet1!A1" not in addrs
        assert "Sheet1!B1" not in addrs

    def test_collect_choose_by_domain_index_prunes_out_of_range_branches(self) -> None:
        """CHOOSE with index in {2}: collects only branch 2 refs."""
        ast = _parse_selector(
            "CHOOSE(Sheet1!X1, Sheet1!A1+Sheet1!B1, Sheet1!C1+Sheet1!D1, Sheet1!E1)"
        )
        env = _make_env(
            {"Sheet1!X1": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({2})))}
        )
        addrs = dynamic_refs_mod._collect_addresses_needing_domain(
            ast, env=env, limits=self._limits()
        )
        # Index X1 is always 2 -> only branch 2 (C1+D1) needed plus X1 itself
        assert "Sheet1!C1" in addrs
        assert "Sheet1!D1" in addrs
        assert "Sheet1!A1" not in addrs
        assert "Sheet1!B1" not in addrs
        assert "Sheet1!E1" not in addrs

    def test_collect_skips_infeasible_then_branch_via_domain_narrowing(self) -> None:
        """IF(A1>5, A1+B1, C1): A1 in [1,3] makes then-branch infeasible → B1 not collected."""
        ast = _parse_selector("IF(Sheet1!A1>5, Sheet1!A1+Sheet1!B1, Sheet1!C1)")
        env = _make_env(
            {
                "Sheet1!A1": CellType(kind=CellKind.NUMBER, interval=IntervalDomain(min=1, max=3)),
            }
        )
        addrs = dynamic_refs_mod._collect_addresses_needing_domain(
            ast, env=env, limits=self._limits()
        )
        # A1>5 narrows A1 to empty in then-branch → then-branch infeasible.
        # B1 only appears in the infeasible then-branch → must not be collected.
        assert "Sheet1!B1" not in addrs
        # C1 (else-branch) is reachable and must be collected.
        assert "Sheet1!C1" in addrs

    def test_collect_without_env_is_backward_compatible(self) -> None:
        """_collect_addresses_needing_domain with no env/limits collects all addresses (old behaviour)."""
        ast = _parse_selector("IF(1, Sheet1!A1+Sheet1!B1, Sheet1!C1+Sheet1!D1)")
        addrs = dynamic_refs_mod._collect_addresses_needing_domain(ast)
        # Without env/limits, fall back to old behaviour: all refs collected
        assert "Sheet1!A1" in addrs
        assert "Sheet1!C1" in addrs


# ---------------------------------------------------------------------------
# Phase 4: Lazy IF evaluation in expr_eval
# ---------------------------------------------------------------------------


class TestLazyIfEval:
    """IF and CHOOSE should not evaluate dead branches in the restricted evaluator."""

    def _eval(self, expr: str, env: dict[str, object] | None = None) -> object:
        from typing import cast

        from excel_grapher.core.expr_eval import evaluate_expr
        from excel_grapher.core.types import CellValue, XlError

        cell_values = env or {}

        def get_cell_value(addr: str) -> CellValue:
            v = cell_values.get(addr)
            if v is None:
                return XlError.REF
            return cast(CellValue, v)

        return evaluate_expr(parse_ast("=" + expr), get_cell_value=get_cell_value)

    def test_if_true_does_not_evaluate_false_branch(self) -> None:
        """IF(TRUE, 1, 1/0) should return 1 without raising or returning an error."""
        from excel_grapher.core.types import XlError

        result = self._eval("IF(TRUE, 1, 1/0)")
        assert result == 1 or result == 1.0, f"Expected 1, got {result!r}"
        assert not isinstance(result, XlError), "Dead branch (1/0) should not be evaluated"

    def test_if_false_does_not_evaluate_true_branch(self) -> None:
        """IF(FALSE, 1/0, 2) should return 2 without raising or returning an error."""
        from excel_grapher.core.types import XlError

        result = self._eval("IF(FALSE, 1/0, 2)")
        assert result == 2 or result == 2.0, f"Expected 2, got {result!r}"
        assert not isinstance(result, XlError), "Dead branch (1/0) should not be evaluated"

    def test_if_false_branch_error_does_not_propagate(self) -> None:
        """IF(1, 42, 1/0) should yield 42, not propagate the division-by-zero error."""
        from excel_grapher.core.types import XlError

        result = self._eval("IF(1, 42, 1/0)")
        assert result == 42 or result == 42.0
        assert not isinstance(result, XlError)

    def test_if_true_branch_error_does_not_propagate(self) -> None:
        """IF(0, 1/0, 99) should yield 99, not propagate the division-by-zero error."""
        from excel_grapher.core.types import XlError

        result = self._eval("IF(0, 1/0, 99)")
        assert result == 99 or result == 99.0
        assert not isinstance(result, XlError)


class TestNumericDomainPrimitives:
    """Shared collapse/densify primitives behind the numeric-domain transfer functions."""

    @staticmethod
    def _limits(max_branches: int = 8) -> DynamicRefLimits:
        return DynamicRefLimits(max_branches=max_branches)

    def test_finite_or_hull_keeps_small_sets_exact(self) -> None:
        out = dynamic_refs_mod._finite_or_hull({3, 1, 2}, self._limits())
        assert out == dynamic_refs_mod._FiniteInts(frozenset({1, 2, 3}))

    def test_finite_or_hull_collapses_oversized_sets_to_bounds(self) -> None:
        out = dynamic_refs_mod._finite_or_hull(set(range(0, 20)), self._limits())
        assert out == dynamic_refs_mod._IntBounds(0, 19)

    def test_finite_or_hull_collapses_at_exact_threshold(self) -> None:
        """`max_branches` is inclusive: a set of exactly that size stays exact."""
        out = dynamic_refs_mod._finite_or_hull(set(range(8)), self._limits(8))
        assert out == dynamic_refs_mod._FiniteInts(frozenset(range(8)))

    def test_densify_or_bounds_expands_narrow_spans(self) -> None:
        out = dynamic_refs_mod._densify_or_bounds(2, 5, self._limits())
        assert out == dynamic_refs_mod._FiniteInts(frozenset({2, 3, 4, 5}))

    def test_densify_or_bounds_keeps_wide_spans_as_bounds(self) -> None:
        out = dynamic_refs_mod._densify_or_bounds(0, 100, self._limits())
        assert out == dynamic_refs_mod._IntBounds(0, 100)

    def test_densify_or_bounds_returns_empty_when_inverted(self) -> None:
        """An inverted span is the empty domain, not a backwards `_IntBounds`."""
        out = dynamic_refs_mod._densify_or_bounds(0, -1, self._limits())
        assert out == dynamic_refs_mod._FiniteInts(frozenset())


class TestNumericDomainTransferInvariants:
    """Behaviour of the binary transfer functions that a shared skeleton must preserve."""

    @staticmethod
    def _limits(max_branches: int = 8) -> DynamicRefLimits:
        return DynamicRefLimits(max_branches=max_branches)

    def test_add_bounds_densifies_narrow_result(self) -> None:
        out = dynamic_refs_mod._add_numeric_domains(
            dynamic_refs_mod._IntBounds(0, 1), dynamic_refs_mod._IntBounds(0, 1), self._limits()
        )
        assert out == dynamic_refs_mod._FiniteInts(frozenset({0, 1, 2}))

    def test_sub_bounds_uses_cross_corners(self) -> None:
        out = dynamic_refs_mod._sub_numeric_domains(
            dynamic_refs_mod._IntBounds(0, 10), dynamic_refs_mod._IntBounds(1, 2), self._limits()
        )
        assert out == dynamic_refs_mod._IntBounds(-2, 9)

    def test_mul_bounds_never_densifies(self) -> None:
        """MUL intentionally skips the densify step that ADD/SUB apply."""
        out = dynamic_refs_mod._mul_numeric_domains(
            dynamic_refs_mod._IntBounds(1, 2), dynamic_refs_mod._IntBounds(1, 2), self._limits()
        )
        assert out == dynamic_refs_mod._IntBounds(1, 4)

    def test_mul_finite_pair_stays_exact(self) -> None:
        out = dynamic_refs_mod._mul_numeric_domains(
            dynamic_refs_mod._FiniteInts(frozenset({2, 3})),
            dynamic_refs_mod._FiniteInts(frozenset({4})),
            self._limits(),
        )
        assert out == dynamic_refs_mod._FiniteInts(frozenset({8, 12}))

    def test_union_merges_finite_sets_without_cartesian_product(self) -> None:
        out = dynamic_refs_mod._union_numeric_domains(
            dynamic_refs_mod._FiniteInts(frozenset({1, 2})),
            dynamic_refs_mod._FiniteInts(frozenset({3})),
            self._limits(),
        )
        assert out == dynamic_refs_mod._FiniteInts(frozenset({1, 2, 3}))

    @pytest.mark.parametrize(
        "fn",
        ["_union_numeric_domains", "_add_numeric_domains", "_sub_numeric_domains"],
    )
    def test_none_operand_propagates(self, fn: str) -> None:
        op = getattr(dynamic_refs_mod, fn)
        assert op(None, dynamic_refs_mod._IntBounds(0, 1), self._limits()) is None
        assert op(dynamic_refs_mod._IntBounds(0, 1), None, self._limits()) is None

    def test_div_returns_none_when_all_divisors_are_zero(self) -> None:
        out = dynamic_refs_mod._div_numeric_domains(
            dynamic_refs_mod._FiniteInts(frozenset({4})),
            dynamic_refs_mod._FiniteInts(frozenset({0})),
            self._limits(),
        )
        assert out is None

    def test_div_truncates_toward_zero(self) -> None:
        out = dynamic_refs_mod._div_numeric_domains(
            dynamic_refs_mod._FiniteInts(frozenset({-7})),
            dynamic_refs_mod._FiniteInts(frozenset({2})),
            self._limits(),
        )
        assert out == dynamic_refs_mod._FiniteInts(frozenset({-3}))

    def test_min_and_max_take_componentwise_bounds(self) -> None:
        a = dynamic_refs_mod._IntBounds(0, 10)
        b = dynamic_refs_mod._IntBounds(20, 30)
        assert dynamic_refs_mod._min_numeric_domains(a, b, self._limits()) == (
            dynamic_refs_mod._IntBounds(0, 10)
        )
        assert dynamic_refs_mod._max_numeric_domains(a, b, self._limits()) == (
            dynamic_refs_mod._IntBounds(20, 30)
        )

    def test_min_of_finite_pair_is_pointwise(self) -> None:
        out = dynamic_refs_mod._min_numeric_domains(
            dynamic_refs_mod._FiniteInts(frozenset({1, 5})),
            dynamic_refs_mod._FiniteInts(frozenset({3})),
            self._limits(),
        )
        assert out == dynamic_refs_mod._FiniteInts(frozenset({1, 3}))
