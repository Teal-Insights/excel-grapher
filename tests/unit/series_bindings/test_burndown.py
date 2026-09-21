"""Tests for internal-binding coverage burndown helpers."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.series_bindings.burndown import (
    collapse_unbound_cells_to_ranges,
    contiguous_column_ranges,
    find_unbound_internal_formula_cells,
    format_burndown_report,
    format_row_column_spans,
    group_unbound_cells_by_sheet_row,
    internal_binding_burndown,
    layout_hint_for_span,
    manifest_binding_addresses,
)
from excel_grapher.series_bindings.workflow import validate_bindings_workbook
from tests.unit.series_bindings.authoring_helpers import (
    public_io_series,
    write_authoring_workbook,
    write_shards,
    years_internal_series,
)


def test_contiguous_column_ranges_groups_gaps() -> None:
    assert contiguous_column_ranges([2, 3, 4, 7, 8]) == [(2, 4), (7, 8)]


def test_group_unbound_cells_by_sheet_row() -> None:
    grouped = group_unbound_cells_by_sheet_row(("Engine!B2", "Engine!C2", "Outputs!B1"))
    assert grouped == {"Engine": {2: [2, 3]}, "Outputs": {1: [2]}}


def test_group_unbound_cells_by_sheet_row_strips_quoted_sheet_names() -> None:
    grouped = group_unbound_cells_by_sheet_row(("'Climate Database'!B26",))
    assert grouped == {"Climate Database": {26: [2]}}


def test_collapse_unbound_cells_to_ranges_joins_consecutive_rows() -> None:
    cells = ("Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3", "Outputs!B1")
    assert collapse_unbound_cells_to_ranges(cells) == (
        "Engine!B2:C3",
        "Outputs!B1",
    )


def test_collapse_unbound_cells_to_ranges_splits_column_gaps() -> None:
    cells = ("Store!B2", "Store!D2", "Store!B3", "Store!D3")
    assert collapse_unbound_cells_to_ranges(cells) == (
        "Store!B2:B3",
        "Store!D2:D3",
    )


def test_format_row_column_spans() -> None:
    assert format_row_column_spans(sheet="Engine", row=2, columns=[2, 3, 4]) == "Engine!B2:D2"
    assert format_row_column_spans(sheet="Engine", row=2, columns=[2, 4]) == "Engine!B2, Engine!D2"


def test_layout_hint_for_span_includes_matrix() -> None:
    assert layout_hint_for_span(width=1, height=1) == "scalar"
    assert layout_hint_for_span(width=2, height=1) == "series"
    assert layout_hint_for_span(width=1, height=4) == "series"
    assert layout_hint_for_span(width=2, height=3) == "matrix"


def test_burndown_groups_unbound_formula_cells(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    result = validate_bindings_workbook(workbook, bindings_dir)
    unbound = find_unbound_internal_formula_cells(
        graph=result["graph"],
        bindings=result["bindings"],
    )
    assert unbound == ("Engine!B2", "Engine!C2")
    grouped = group_unbound_cells_by_sheet_row(unbound)
    assert grouped == {"Engine": {2: [2, 3]}}
    assert (
        format_row_column_spans(sheet="Engine", row=2, columns=grouped["Engine"][2])
        == "Engine!B2:C2"
    )
    assert collapse_unbound_cells_to_ranges(unbound) == ("Engine!B2:C2",)

    report = internal_binding_burndown(result["graph"], result["bindings"])
    assert report.unbound_count == 2
    assert report.collapsed_ranges == ("Engine!B2:C2",)
    assert report.collapsed_layout_hints == ("series",)
    assert report.formula_node_count >= 4


def test_burndown_reports_no_unbound_cells_when_internals_cover_formulas(
    tmp_path: Path,
) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[years_internal_series(fill=True)],
    )
    result = validate_bindings_workbook(workbook, bindings_dir)
    report = internal_binding_burndown(result["graph"], result["bindings"])
    assert report.unbound_cells == ()
    assert report.unbound_count == 0


def test_burndown_honors_exempt_addresses(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    result = validate_bindings_workbook(workbook, bindings_dir)
    report = internal_binding_burndown(
        result["graph"],
        result["bindings"],
        exempt_cells=frozenset({"Engine!B2", "Engine!C2"}),
    )
    assert report.unbound_cells == ()


def test_manifest_binding_addresses_expands_list_data_range(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx")
    inputs, result_a, result_b = public_io_series()
    internal = years_internal_series(data_range="Engine!B2")
    internal["data_range"] = ["Engine!B2", "Engine!C2"]
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
        internals=[internal],
    )
    result = validate_bindings_workbook(workbook, bindings_dir)
    addresses = manifest_binding_addresses(
        result["graph"],
        result["bindings"],
        direction="internal",
    )
    assert "Engine!B2" in addresses
    assert "Engine!C2" in addresses


def test_collapse_two_dimensional_block_hints_matrix() -> None:
    cells = ("Engine!B2", "Engine!C2", "Engine!B3", "Engine!C3")
    ranges = collapse_unbound_cells_to_ranges(cells)
    assert ranges == ("Engine!B2:C3",)
    from excel_grapher.series_bindings.burndown import layout_hint_for_range

    assert layout_hint_for_range(ranges[0]) == "matrix"


def test_burndown_max_rows_marks_truncation(tmp_path: Path) -> None:
    workbook = write_authoring_workbook(tmp_path / "workbook.xlsx", extra_engine_row=True)
    inputs, result_a, result_b = public_io_series()
    bindings_dir = write_shards(
        tmp_path / "bindings",
        inputs=[inputs],
        outputs=[result_a, result_b],
    )
    result = validate_bindings_workbook(workbook, bindings_dir)
    report = internal_binding_burndown(result["graph"], result["bindings"], max_rows=0)
    assert report.rows_truncated is True
    rendered = "\n".join(format_burndown_report(report))
    assert "... (truncated)" in rendered
    assert "bound graph closure" in rendered.lower() or "bound closure" in rendered.lower()
