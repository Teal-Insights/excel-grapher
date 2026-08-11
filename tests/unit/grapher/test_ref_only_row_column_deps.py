"""ROW/COLUMN/ROWS/COLUMNS address-only refs must not create dependency edges (#515)."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher import create_dependency_graph
from excel_grapher.core.cell_types import CellKind, CellType, EnumDomain
from excel_grapher.grapher.dependency_provenance import DependencyCause
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefLimits
from excel_grapher.grapher.parser import (
    mask_ref_only_function_calls,
    parse_range_refs_with_spans,
    parse_standalone_cell_refs,
    ref_only_function_spans,
)


def test_ref_only_function_spans_masks_simple_row_column_calls() -> None:
    formula = "=ROW($B$1)+COLUMN($C$2)+A1"
    spans = ref_only_function_spans(formula)
    masked = mask_ref_only_function_calls(formula)
    assert len(spans) == 2
    refs = parse_standalone_cell_refs(masked)
    assert {(r.column, r.row) for r in refs} == {("A", 1)}


def test_ref_only_function_spans_masks_rows_columns_ranges() -> None:
    formula = "=ROWS($A$1:$A$2)+COLUMNS($B$1:$C$1)+D1"
    masked = mask_ref_only_function_calls(formula)
    assert parse_range_refs_with_spans(masked) == []
    refs = parse_standalone_cell_refs(masked)
    assert {(r.column, r.row) for r in refs} == {("D", 1)}


def test_ref_only_function_spans_keeps_nested_index_refs() -> None:
    formula = "=ROW(INDEX(Sheet1!A1:A3,Sheet1!B1))"
    # Nested call: do not mask the outer ROW(...); INDEX refs stay visible.
    assert ref_only_function_spans(formula) == []
    masked = mask_ref_only_function_calls(formula)
    assert "A1" in masked and "B1" in masked


def test_row_column_rows_columns_address_only_refs_are_not_deps(tmp_path: Path) -> None:
    excel_path = tmp_path / "row_col_ref_only.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 42)  # B1
    ws.write_number(0, 0, 1)  # A1
    ws.write_number(1, 0, 2)  # A2
    ws.write_formula(0, 2, "=ROW($B$1)", None, 1)  # C1
    ws.write_formula(1, 2, "=COLUMN($B$1)", None, 2)  # C2
    ws.write_formula(2, 2, "=ROW($B$1)+COLUMN($B$1)", None, 3)  # C3
    ws.write_formula(0, 3, "=ROWS($A$1:$A$2)", None, 2)  # D1
    ws.write_formula(1, 3, "=COLUMNS($A$1:$A$2)", None, 1)  # D2
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!C1", "Sheet1!C2", "Sheet1!C3", "Sheet1!D1", "Sheet1!D2"],
        load_values=False,
        capture_dependency_provenance=True,
    )

    for key in ["Sheet1!C1", "Sheet1!C2", "Sheet1!C3", "Sheet1!D1", "Sheet1!D2"]:
        assert graph.get_dependencies(key) == frozenset(), key
        assert not any(src == key for src, _dst in graph._edge_provenance)


def test_ref_only_same_address_in_value_position_keeps_edge(tmp_path: Path) -> None:
    excel_path = tmp_path / "row_and_value.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 1, 5)  # B1
    ws.write_number(0, 2, 7)  # C1
    ws.write_formula(0, 0, "=ROW($B$1)+$B$1", None, 6)  # A1
    ws.write_formula(1, 0, "=C1+ROW($B$1)", None, 7)  # A2
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!A1", "Sheet1!A2"],
        load_values=False,
        capture_dependency_provenance=True,
    )
    assert graph.get_dependencies("Sheet1!A1") == frozenset({"Sheet1!B1"})
    assert graph.get_dependencies("Sheet1!A2") == frozenset({"Sheet1!C1"})
    prov = graph.get_edge_attrs("Sheet1!A1", "Sheet1!B1").provenance
    assert prov is not None
    assert DependencyCause.direct_ref in prov.causes


def test_row_anchor_inside_offset_is_not_self_loop(tmp_path: Path) -> None:
    """LIC-DSF pattern: ROW($B$106) on B106 must not create B106→B106 (#515)."""
    excel_path = tmp_path / "row_anchor_offset.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number(0, 0, 1)  # A1 OFFSET base
    # B106 (row 106, 0-indexed 105): row offset is address-only via ROW($B$106).
    ws.write_formula(
        105,
        1,
        "=OFFSET(Sheet1!$A$1,ROW()-ROW($B$106),0)",
        None,
        1,
    )
    wb.close()

    config = DynamicRefConfig(cell_type_env={}, limits=DynamicRefLimits())
    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!B106"],
        load_values=False,
        dynamic_refs=config,
        capture_dependency_provenance=True,
    )
    deps = graph.get_dependencies("Sheet1!B106")
    assert "Sheet1!B106" not in deps
    assert "Sheet1!A1" in deps
    report = graph.cycle_report()
    assert not any("Sheet1!B106" in c and len(c) == 1 for c in report.must_cycles)


def test_row_anchor_inside_offset_index_is_not_self_loop(tmp_path: Path) -> None:
    """Same anchor pattern with INDEX base (issue #515 LIC-DSF shape)."""
    excel_path = tmp_path / "row_anchor_offset_index.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    for r in range(5):
        ws.write_number(r, 0, r + 1)  # A1:A5
    ws.write_formula(
        105,
        1,
        "=OFFSET(INDEX(Sheet1!$A$1:$A$5,ROW()-ROW($B$106)+1,1),0,0)",
        None,
        1,
    )
    wb.close()

    env = {
        f"Sheet1!A{r}": CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({r})))
        for r in range(1, 6)
    }
    config = DynamicRefConfig(cell_type_env=env, limits=DynamicRefLimits())
    graph = create_dependency_graph(
        excel_path,
        ["Sheet1!B106"],
        load_values=False,
        dynamic_refs=config,
    )
    deps = graph.get_dependencies("Sheet1!B106")
    assert "Sheet1!B106" not in deps
    assert "Sheet1!A1" in deps
    assert not any("Sheet1!B106" in c and len(c) == 1 for c in graph.cycle_report().must_cycles)
