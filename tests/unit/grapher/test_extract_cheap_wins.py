"""Ops-count regressions for issue #715 extract-path cheap wins.

Wall-clock is not the oracle: these tests count full-env scans, named-range
regex compiles, and per-formula workbook bounds walks.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest
import xlsxwriter

from excel_grapher.core.formula_ast import parse_preserving_axes
from excel_grapher.core.formula_normalization import (
    build_named_range_replacement_state,
    expand_defined_names,
)
from excel_grapher.grapher.builder import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefError, DynamicRefLimits


def test_named_range_state_compile_is_skipped_when_state_is_passed() -> None:
    named_ranges = {f"Name{i}": ("Sheet1", f"A{i}") for i in range(1, 95)}
    state = build_named_range_replacement_state(named_ranges, {})
    compile_ops = 0
    original = build_named_range_replacement_state

    def counting(
        named_ranges: dict[str, tuple[str, str]] | None = None,
        named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
    ) -> object:
        nonlocal compile_ops
        compile_ops += 1
        return original(named_ranges, named_range_ranges)

    with patch(
        "excel_grapher.core.formula_normalization.build_named_range_replacement_state",
        counting,
    ):
        for _ in range(50):
            expanded = expand_defined_names(
                "=A1+Name1",
                named_ranges=named_ranges,
                named_range_ranges={},
                name_state=state,
            )
            assert "A1" in expanded

    assert compile_ops == 0


def test_parse_preserving_axes_name_state_matches_named_ranges_map() -> None:
    named_ranges = {"MyName": ("Sheet1", "Z9")}
    state = build_named_range_replacement_state(named_ranges, {})
    with_state = parse_preserving_axes(
        "=MyName+A1",
        anchor="Sheet1!B2",
        name_state=state,
    )
    with_maps = parse_preserving_axes(
        "=MyName+A1",
        anchor="Sheet1!B2",
        named_ranges=named_ranges,
    )
    assert with_state == with_maps


def test_create_dependency_graph_compiles_named_range_state_once(tmp_path: Path) -> None:
    excel_path = tmp_path / "named.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)
    wb.define_name("Rate", "=Sheet1!$A$1")
    for row in range(1, 41):
        ws.write_formula(row, 1, "=Rate+1")
    wb.close()

    compile_ops = 0
    original = build_named_range_replacement_state

    def counting(
        named_ranges: dict[str, tuple[str, str]] | None = None,
        named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
    ) -> object:
        nonlocal compile_ops
        compile_ops += 1
        return original(named_ranges, named_range_ranges)

    with (
        patch(
            "excel_grapher.core.formula_normalization.build_named_range_replacement_state",
            counting,
        ),
        patch(
            "excel_grapher.grapher.parser.build_named_range_replacement_state",
            counting,
        ),
    ):
        graph = create_dependency_graph(
            excel_path,
            [f"Sheet1!B{row}" for row in range(2, 42)],
            load_values=False,
        )

    assert "Sheet1!B2" in graph
    assert "Sheet1!A1" in graph.get_dependencies("Sheet1!B2")
    assert compile_ops == 1


def test_provenance_does_not_rescan_workbook_sheet_bounds_per_formula(tmp_path: Path) -> None:
    excel_path = tmp_path / "ranges.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    for title in ("Alpha", "Beta", "Gamma"):
        ws = wb.add_worksheet(title)
        ws.write_number(0, 0, 1)
        for row in range(1, 21):
            ws.write_formula(row, 0, f"=SUM({title}!A$1:A$1)")
    wb.close()

    import excel_grapher.grapher.provenance_collect as provenance_collect

    bounds_ops = 0
    original = provenance_collect._sheet_bounds_from_workbook

    def counting(wb_formulas: object) -> dict[str, tuple[int, int]]:
        nonlocal bounds_ops
        bounds_ops += 1
        return original(wb_formulas)

    with patch.object(provenance_collect, "_sheet_bounds_from_workbook", counting):
        graph = create_dependency_graph(
            excel_path,
            ["Alpha!A21", "Beta!A21", "Gamma!A21"],
            load_values=False,
            capture_dependency_provenance=True,
            expand_ranges=True,
        )

    assert "Alpha!A21" in graph
    assert bounds_ops <= 1


def test_index_missing_leaf_constraint_still_fail_closes(tmp_path: Path) -> None:
    excel_path = tmp_path / "index_missing.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 10)
    ws.write_number(1, 0, 20)
    ws.write_number(2, 0, 30)
    ws.write_number(0, 1, 2)
    ws.write_formula(0, 2, "=INDEX(Sheet1!A1:A3,Sheet1!B1)", None, 20)
    wb.close()

    with pytest.raises(DynamicRefError, match="no constraint"):
        create_dependency_graph(
            excel_path,
            ["Sheet1!C1"],
            load_values=False,
            dynamic_refs=DynamicRefConfig(cell_type_env={}, limits=DynamicRefLimits()),
        )
