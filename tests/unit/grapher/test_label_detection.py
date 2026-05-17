"""Unit tests for ``excel_grapher.grapher.label_detection``."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.blank_ranges import parse_blank_range_spec
from excel_grapher.grapher.label_detection import (
    BehaviorRule,
    LabelDetectionConfig,
    LabelDetectionState,
    RegionLabelParams,
    RegionSelector,
    RegionSpec,
    build_label_behavior_registry,
    collect_labels_for_node,
    region_specs_from_ranges,
    selector_matches,
)


def test_region_specs_from_ranges_matches_blank_range_parse() -> None:
    specs = region_specs_from_ranges(["Sheet1!B2:D4", "'A B'!A1"])
    assert len(specs) == 2
    r0 = parse_blank_range_spec("Sheet1!B2:D4")
    assert (
        specs[0].sheet,
        specs[0].min_row,
        specs[0].max_row,
        specs[0].min_col,
        specs[0].max_col,
    ) == (
        r0[0],
        r0[1],
        r0[3],
        r0[2],
        r0[4],
    )


def test_region_specs_from_ranges_rejects_unqualified() -> None:
    with pytest.raises(ValueError, match="sheet-qualified"):
        region_specs_from_ranges(["B2:D4"])


def test_selector_matches_include_exclude() -> None:
    inc = region_specs_from_ranges(["Sheet1!A1:C10"])
    exc = region_specs_from_ranges(["Sheet1!B2:B2"])
    sel = RegionSelector(include=inc, exclude=exc)
    assert selector_matches("Sheet1", 2, 1, sel) is True
    assert selector_matches("Sheet1", 2, 2, sel) is False
    assert selector_matches("Other", 2, 1, sel) is False


def test_selector_empty_include_never_matches() -> None:
    sel = RegionSelector(include=(), exclude=())
    assert selector_matches("Sheet1", 1, 1, sel) is False


def test_collect_labels_unknown_behavior_raises() -> None:
    cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("no_such_behavior",))
    reg = build_label_behavior_registry(None)
    st = LabelDetectionState()
    with pytest.raises(ValueError, match="Unknown label detection behavior"):
        collect_labels_for_node(
            key="S!A1",
            sheet="S",
            row=1,
            col=1,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=None,
            ws_formulas=None,
        )


def test_collect_labels_rule_stop_skips_fallback() -> None:
    cfg = LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="emptyRegion",
                selector=RegionSelector(include=(RegionSpec("S", 1, 5, 1, 5),)),
                behaviors=(),
                stop_after_match=True,
            ),
        ),
        fallback_behaviors=("left_edge_scan",),
    )
    reg = build_label_behavior_registry(None)
    st = LabelDetectionState()
    row, col = collect_labels_for_node(
        key="S!C1",
        sheet="S",
        row=1,
        col=3,
        cfg=cfg,
        registry=reg,
        state=st,
        ws_values=None,
        ws_formulas=None,
    )
    assert row == [] and col == []


def test_collect_labels_region_label_columns(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "label_wb.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "RowTitle")
    ws.write_number(0, 2, 42)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(
            enabled=True,
            rules=(
                BehaviorRule(
                    name="r",
                    selector=RegionSelector(include=(RegionSpec("S", 1, 3, 1, 5),)),
                    behaviors=("region_left_label_columns",),
                    region_params=RegionLabelParams(label_columns=("A",)),
                ),
            ),
            fallback_behaviors=(),
        )
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!C1",
            sheet="S",
            row=1,
            col=3,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["RowTitle"]
        assert col == []
    finally:
        wv.close()


def test_left_edge_scan_skips_leading_non_year_numbers(tmp_path) -> None:
    """Non-year numbers break only after at least one label was collected to the right."""
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "skip_num.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "A-label")
    ws.write_number(0, 1, 999)
    ws.write_string(0, 2, "C-label")
    ws.write_number(0, 3, 888)
    ws.write_number(0, 4, 1)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True)
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!E1",
            sheet="S",
            row=1,
            col=5,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["C-label"]
        assert col == []
    finally:
        wv.close()


def test_top_edge_scan_skips_leading_non_year_numbers(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "skip_num_col.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "top")
    ws.write_number(1, 0, 999)
    ws.write_string(2, 0, "mid")
    ws.write_number(3, 0, 888)
    ws.write_number(4, 0, 1)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True)
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!A5",
            sheet="S",
            row=5,
            col=1,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == []
        assert col == ["mid"]
    finally:
        wv.close()
