"""
Tests end-to-end user workflows for label detection, including default
behaviors and custom behavior registry.

These tests are intended to map to the micro-workbook examples in
examples/micro_workbooks/label_detection.qmd, but without dependency on
or logical/semantic coupling to the xlsx file. Unimplemented roadmap
behaviors are marked ``xfail``.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import column_index_from_string

from excel_grapher.grapher import (
    BehaviorRule,
    LabelDetectionBehavior,
    LabelDetectionConfig,
    LabelDetectionContext,
    LabelResult,
    RegionLabelParams,
    RegionSelector,
    create_dependency_graph,
    region_specs_from_ranges,
)
from tests.integration.user_flows.utils import WorkbookFactory, build_workbook_factory


@pytest.fixture
def label_workbook_factory(tmp_path: Path) -> WorkbookFactory:
    return build_workbook_factory(tmp_path, prefix="label_detection")


def _labels(
    workbook: Path,
    target: str,
    *,
    label_detection: LabelDetectionConfig | None = None,
    label_behaviors: list[LabelDetectionBehavior] | None = None,
) -> dict[str, Any]:
    cfg = label_detection or LabelDetectionConfig(enabled=True)
    graph = create_dependency_graph(
        workbook,
        [target],
        load_values=True,
        label_detection=cfg,
        label_behaviors=label_behaviors,
    )
    node = graph.get_node(target)
    assert node is not None
    return dict(node.metadata)


def _row_labels(metadata: dict[str, Any]) -> list[str]:
    return list(metadata.get("row_labels", []))


def _column_labels(metadata: dict[str, Any]) -> list[str]:
    return list(metadata.get("column_labels", []))


# --- Neighboring row/column labels ---


def test_label_detection_disabled_leaves_metadata_empty(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("B1", "Column 1")
        ws.write("A2", "Row 1")
        ws.write_number("B2", 0)

    path = label_workbook_factory(_populate)
    metadata = _labels(
        path,
        "Sheet1!B2",
        label_detection=LabelDetectionConfig(enabled=False),
    )
    assert _row_labels(metadata) == []
    assert _column_labels(metadata) == []


def test_neighbor_strings_become_row_and_column_labels(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("B1", "Column 1")
        ws.write("A2", "Row 1")
        ws.write_number("B2", 0)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!B2")
    assert _row_labels(metadata) == ["Row 1"]
    assert _column_labels(metadata) == ["Column 1"]


# --- Non-neighboring row/column labels ---


def test_scan_skips_numeric_cells_to_find_string_labels(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("B1", "First Column")
        ws.write("C1", "Second Column")
        ws.write("A2", "First Row")
        ws.write_number("B2", 1)
        ws.write_number("C2", 0)
        ws.write("A3", "Second Row")
        ws.write_number("B3", 0)
        ws.write_number("C3", 1)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!C3")
    assert _row_labels(metadata) == ["Second Row"]
    assert _column_labels(metadata) == ["Second Column"]


# --- No labels ---


def test_blank_cells_stop_scan_with_no_labels(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write_number("B1", 0)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!B1")
    assert _row_labels(metadata) == []
    assert _column_labels(metadata) == []


# --- Nested row/column labels ---


def test_nested_header_rows_and_columns_are_collected(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("C1", "Column 1.1")
        ws.write("D1", "Column 2.1")
        ws.write("C2", "Column 1.2")
        ws.write("D2", "Column 2.2")
        ws.write("A3", "Row 1.1")
        ws.write("B3", "Row 1.2")
        ws.write_number("C3", 0)
        ws.write_number("D3", 0)
        ws.write("A4", "Row 1.2")
        ws.write("B4", "Row  2.2")
        ws.write_number("C4", 0)
        ws.write_number("D4", 0)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!D4")
    assert _row_labels(metadata) == ["Row  2.2", "Row 1.2"]
    assert _column_labels(metadata) == ["Column 2.2", "Column 2.1"]


# --- Merged cells ---


def test_merged_cell_label_is_included_in_row_labels(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("C1", "Column 1.1")
        ws.write("C2", "Column 1.2")
        ws.write("D2", "Column 2.2")
        ws.merge_range("A3:A4", "Row 1.1")
        ws.write("B3", "Row 1.2")
        ws.write_number("C3", 0)
        ws.write_number("D3", 0)
        ws.write("B4", "Row  2.2")
        ws.write_number("C4", 0)
        ws.write_number("D4", 0)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!D4")
    assert "Row 1.1" in _row_labels(metadata)


# --- Intervening blank cells ---


def test_blank_gap_prevents_labels_from_other_tables(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("C1", "List")
        ws.write("D1", "List Item 1")
        ws.write("E1", "List Item 2")
        ws.write("A3", "List")
        ws.write("D3", "Column_1")
        ws.write("E3", "Column_2")
        ws.write("A4", "List Item 1")
        ws.write("C4", "Row_1")
        ws.write_number("D4", 1)
        ws.write_number("E4", 2)
        ws.write("A5", "List Item 2")
        ws.write("C5", "Row_2")
        ws.write_number("D5", 1)
        ws.write_number("E5", 4 / 3)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!E5")
    assert _row_labels(metadata) == ["Row_2"]
    assert _column_labels(metadata) == ["Column_2"]


# --- Intervening text cells ---


def test_tall_format_gdp_row_label_is_year_only(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("B1", "Population")
        ws.write("C1", "Country")
        ws.write("D1", "GDP")
        ws.write("A2", "Year 1")
        ws.write_number("B2", 349)
        ws.write("C2", "USA")
        ws.write_number("D2", 31_680_000)
        ws.write("B4", "Year 1")
        ws.write("A5", "Population")
        ws.write_number("B5", 349)
        ws.write("A6", "Country")
        ws.write("B6", "USA")
        ws.write("A7", "GDP")
        ws.write_number("B7", 31_680)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!D2", label_detection=LabelDetectionConfig(enabled=True))
    assert _row_labels(metadata) == ["Year 1"]
    assert "USA" not in _row_labels(metadata)


def test_wide_format_gdp_column_label_is_year_only(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("B1", "Population")
        ws.write("C1", "Country")
        ws.write("D1", "GDP")
        ws.write("A2", "Year 1")
        ws.write_number("B2", 349)
        ws.write("C2", "USA")
        ws.write_number("D2", 31_680_000)
        ws.write("B4", "Year 1")
        ws.write("A5", "Population")
        ws.write_number("B5", 349)
        ws.write("A6", "Country")
        ws.write("B6", "USA")
        ws.write("A7", "GDP")
        ws.write_number("B7", 31_680)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!B7", label_detection=LabelDetectionConfig(enabled=True))
    assert _column_labels(metadata) == ["Year 1"]
    assert "USA" not in _column_labels(metadata)


def test_wide_format_gdp_collects_identifier_and_year_with_full_scans(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("B1", "Population")
        ws.write("C1", "Country")
        ws.write("D1", "GDP")
        ws.write("A2", "Year 1")
        ws.write_number("B2", 349)
        ws.write("C2", "USA")
        ws.write_number("D2", 31_680_000)
        ws.write("B4", "Year 1")
        ws.write("A5", "Population")
        ws.write_number("B5", 349)
        ws.write("A6", "Country")
        ws.write("B6", "USA")
        ws.write("A7", "GDP")
        ws.write_number("B7", 31_680)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="wideGdpBlock",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A4:B7"]),
                ),
                behaviors=("full_row_scan", "full_column_scan"),
                stop_after_match=True,
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(path, "Sheet1!B7", label_detection=cfg)
    assert _row_labels(metadata) == ["GDP"]
    assert _column_labels(metadata) == ["USA", "Year 1"]


# --- Duplicate labels ---


def test_duplicate_labels_are_deduplicated(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("C1", "Dupe Label")
        ws.write("C2", "Dupe Label")
        ws.write("A3", "Dupe Label")
        ws.write("B3", "Dupe Label")
        ws.write_number("C3", 0)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!C3")
    assert _row_labels(metadata) == ["Dupe Label"]
    assert _column_labels(metadata) == ["Dupe Label"]


# --- Year labels ---


def test_calendar_year_in_label_column_is_collected(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("A1", "Year")
        ws.write("B1", "Revenue")
        ws.write_number("A2", 1999)
        ws.write_number("B2", 100)
        ws.write_number("A3", 2000)
        ws.write_number("B3", 200)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!B3")
    assert _row_labels(metadata) == ["2000"]
    assert _column_labels(metadata) == ["Revenue"]


# --- Numeric labels ---


def test_non_calendar_numeric_row_id_is_not_collected_by_default(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("A1", "Year")
        ws.write("B1", "Revenue")
        ws.write_number("A2", 1)
        ws.write_number("B2", 100)
        ws.write_number("A3", 2)
        ws.write_number("B3", 200)

    path = label_workbook_factory(_populate)
    metadata = _labels(path, "Sheet1!B3")
    assert _row_labels(metadata) == []
    assert _column_labels(metadata) == ["Revenue"]


@dataclass
class _ColumnARowLabel(LabelDetectionBehavior):
    name: str = "column_a_row_label"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        value = ctx.ws_values.cell(row=ctx.row, column=1).value
        if value is None:
            return LabelResult()
        text = str(value).strip()
        if not text:
            return LabelResult()
        return LabelResult(row_labels=(text,))


def test_custom_behavior_collects_column_a_numeric_row_id(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("A1", "Year")
        ws.write("B1", "Revenue")
        ws.write_number("A2", 1)
        ws.write_number("B2", 100)
        ws.write_number("A3", 2)
        ws.write_number("B3", 200)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="numericRowIds",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A1:B3"]),
                ),
                behaviors=("column_a_row_label", "top_edge_scan"),
                stop_after_match=True,
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(
        path,
        "Sheet1!B3",
        label_detection=cfg,
        label_behaviors=[_ColumnARowLabel()],
    )
    assert _row_labels(metadata) == ["2"]
    assert _column_labels(metadata) == ["Revenue"]


# --- Year-to-offset transform ---


@dataclass
class _YearOffsetRowLabel(LabelDetectionBehavior):
    name: str = "year_offset_row_label"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        region_params = ctx.region_params
        if (
            region_params is None
            or ctx.ws_values is None
            or region_params.min_row is None
            or region_params.max_row is None
        ):
            return LabelResult()

        base_year: int | None = None
        for row in range(region_params.min_row, region_params.max_row + 1):
            value = ctx.ws_values.cell(row=row, column=1).value
            if isinstance(value, int) and 1900 <= value <= 2100:
                base_year = value
                break
        if base_year is None:
            return LabelResult()

        current = ctx.ws_values.cell(row=ctx.row, column=1).value
        if not isinstance(current, int) or not (1900 <= current <= 2100):
            return LabelResult()

        return LabelResult(row_labels=(f"offset:{current - base_year}",))


def test_custom_year_offset_row_label(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("A1", "Year")
        ws.write("B1", "Revenue")
        ws.write_number("A2", 1999)
        ws.write_number("B2", 100)
        ws.write_number("A3", 2000)
        ws.write_number("B3", 200)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="tallYearOffsets",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A1:B3"]),
                ),
                behaviors=("year_offset_row_label", "top_edge_scan"),
                stop_after_match=True,
                region_params=RegionLabelParams(min_row=2, max_row=3),
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(
        path,
        "Sheet1!B3",
        label_detection=cfg,
        label_behaviors=[_YearOffsetRowLabel()],
    )
    assert _row_labels(metadata) == ["offset:1"]
    assert _column_labels(metadata) == ["Revenue"]


# --- Rightward/downward scans ---


def test_right_and_bottom_scans_collect_units_and_source(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write("A1", "Value")
        ws.write("B1", "Multiplier")
        ws.write("C1", "Unit")
        ws.write_number("A2", 100)
        ws.write("B2", "million")
        ws.write("C2", "dollars")
        ws.write_number("A3", 200)
        ws.write("B3", "million")
        ws.write("C3", "dollars")
        ws.write("A4", "Source: CIA Factbook, 2012")

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="unitsAndSourceBlock",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A1:C4"]),
                ),
                behaviors=("right_edge_scan", "bottom_edge_scan", "top_edge_scan"),
                stop_after_match=True,
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(path, "Sheet1!A3", label_detection=cfg)
    assert _column_labels(metadata) == ["Value"]
    assert "million" in _row_labels(metadata)
    assert any("Source" in label for label in _row_labels(metadata))


# --- Left-then-up scans ---


def test_left_then_up_scan_collects_indent_hierarchy(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, wb) -> None:
        bold = wb.add_format({"bold": True})
        indent_1 = wb.add_format({"indent": 1})
        indent_2 = wb.add_format({"indent": 2})
        ws.write("A1", "Country Details")
        ws.write_number("B2", 2010)
        ws.write("A3", "United States", bold)
        ws.write("A4", "GDP ($)", indent_1)
        ws.write("A5", "Nominal", indent_2)
        ws.write_number("B5", 15.01)
        ws.write("A6", "Real", indent_2)
        ws.write_number("B6", 15.31)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="indentHierarchy",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A1:B6"]),
                ),
                behaviors=("left_then_up_scan", "top_edge_scan"),
            ),
        ),
    )
    metadata = _labels(path, "Sheet1!B6", label_detection=cfg)
    assert _row_labels(metadata) == ["United States", "GDP", "Real"]


def test_left_then_up_scan_prioritizes_indent_then_style_rank(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, wb) -> None:
        bold = wb.add_format({"bold": True, "indent": 2})
        italic = wb.add_format({"italic": True, "indent": 2})
        normal_indent_2 = wb.add_format({"indent": 2})
        indent_1 = wb.add_format({"indent": 1})
        ws.write("A2", "United States")
        ws.write("A3", "GDP", indent_1)
        ws.write("A4", "Tier Bold", bold)
        ws.write("A5", "Tier Italic", italic)
        ws.write("A6", "Tier Normal", normal_indent_2)
        ws.write_number("B6", 15.31)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="leftThenUpHierarchy",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A2:B6"]),
                ),
                behaviors=("left_then_up_scan", "top_edge_scan"),
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(path, "Sheet1!B6", label_detection=cfg)
    assert _row_labels(metadata) == [
        "United States",
        "GDP",
        "Tier Bold",
        "Tier Italic",
        "Tier Normal",
    ]


def test_left_then_up_scan_stops_when_no_left_label_column(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, _wb) -> None:
        ws.write_number("B1", 100)
        ws.write_number("C1", 200)
        ws.write_number("D1", 300)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="leftThenUpNoLeftLabelColumn",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!B1:D1"]),
                ),
                behaviors=("left_then_up_scan",),
                stop_after_match=True,
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(path, "Sheet1!D1", label_detection=cfg)
    assert _row_labels(metadata) == []
    assert _column_labels(metadata) == []


@dataclass
class _FontWeightHierarchyRowLabels(LabelDetectionBehavior):
    name: str = "font_weight_hierarchy_row_labels"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        region_params = ctx.region_params
        if (
            ctx.ws_values is None
            or region_params is None
            or not region_params.label_columns
            or region_params.min_row is None
            or region_params.max_row is None
        ):
            return LabelResult()

        col_idx = column_index_from_string(region_params.label_columns[0])
        current_parent: str | None = None
        row_labels: list[str] = []
        for row in range(region_params.min_row, min(ctx.row, region_params.max_row) + 1):
            cell = ctx.ws_values.cell(row=row, column=col_idx)
            if not isinstance(cell.value, str):
                continue
            text = cell.value.strip()
            if not text:
                continue

            is_bold = bool(cell.font and cell.font.bold)
            if is_bold:
                current_parent = text

            if row == ctx.row:
                if is_bold:
                    row_labels = [text]
                elif current_parent is not None:
                    row_labels = [current_parent, text]
                else:
                    row_labels = [text]

        return LabelResult(row_labels=tuple(row_labels))


def test_custom_font_weight_hierarchy_row_labels(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, wb) -> None:
        bold = wb.add_format({"bold": True})
        indent_1 = wb.add_format({"indent": 1})
        indent_2 = wb.add_format({"indent": 2})
        ws.write("A1", "Country Details")
        ws.write_number("B2", 2010)
        ws.write("A3", "United States", bold)
        ws.write("A4", "GDP ($)", indent_1)
        ws.write("A5", "Nominal", indent_2)
        ws.write_number("B5", 15.01)
        ws.write("A6", "Real", indent_2)
        ws.write_number("B6", 15.31)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="fontWeightHierarchyDemo",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A1:B6"]),
                ),
                behaviors=("font_weight_hierarchy_row_labels", "top_edge_scan"),
                stop_after_match=True,
                region_params=RegionLabelParams(
                    label_columns=("A",),
                    min_row=1,
                    max_row=6,
                ),
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(
        path,
        "Sheet1!B6",
        label_detection=cfg,
        label_behaviors=[_FontWeightHierarchyRowLabels()],
    )
    assert _row_labels(metadata) == ["United States", "Real"]
    assert _column_labels(metadata) == []


def test_left_then_up_scan_parents_year_leaf_when_header_is_bold(
    label_workbook_factory: WorkbookFactory,
) -> None:
    def _populate(ws, wb) -> None:
        bold = wb.add_format({"bold": True})
        ws.write("A1", "Year", bold)
        ws.write("B1", "Revenue")
        ws.write_number("A2", 1999)
        ws.write_number("B2", 100)
        ws.write_number("A3", 2000)
        ws.write_number("B3", 200)

    path = label_workbook_factory(_populate)
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="fontWeightYearTable",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A1:B3"]),
                ),
                behaviors=("left_then_up_scan",),
                stop_after_match=True,
                region_params=RegionLabelParams(
                    label_columns=("A",),
                    min_row=1,
                    max_row=3,
                ),
            ),
        ),
        fallback_behaviors=(),
    )
    metadata = _labels(
        path,
        "Sheet1!B3",
        label_detection=cfg,
    )
    assert _row_labels(metadata) == ["Year", "2000"]
    assert _column_labels(metadata) == []
