"""Unit tests for ``excel_grapher.grapher.label_detection``."""

from __future__ import annotations

import pytest

from excel_grapher.grapher.blank_ranges import parse_blank_range_spec
from excel_grapher.grapher.label_detection import (
    BehaviorRule,
    LabelDetectionConfig,
    LabelDetectionState,
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


def test_collect_labels_rejects_year_offset_headers_behavior_name() -> None:
    cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("year_offset_headers",))
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


def test_collect_labels_left_edge_then_up_scan(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "label_wb.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "RowTitle")
    ws.write_string(0, 1, "RowLeaf")
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
                    behaviors=("left_edge_then_up_scan",),
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


def test_left_edge_then_up_scan_same_indent_same_style_is_skipped(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "left_then_up_skip_same_style.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    i2 = wb.add_format({"indent": 2})
    i1 = wb.add_format({"indent": 1})
    ws.write_string(0, 0, "Country")
    ws.write_string(1, 0, "Category", i1)
    ws.write_string(2, 0, "Current", i2)
    ws.write_string(3, 0, "Target", i2)
    ws.write_number(3, 1, 1)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("left_edge_then_up_scan",))
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!B4",
            sheet="S",
            row=4,
            col=2,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["Country", "Category", "Target"]
        assert col == []
    finally:
        wv.close()


def test_left_edge_then_up_scan_stops_on_same_indent_lower_style(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "left_then_up_stop_lower_style.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    bold_i2 = wb.add_format({"bold": True, "indent": 2})
    italic_i2 = wb.add_format({"italic": True, "indent": 2})
    i1 = wb.add_format({"indent": 1})
    ws.write_string(0, 0, "Country")
    ws.write_string(1, 0, "Category", i1)
    ws.write_string(2, 0, "Upper Italic", italic_i2)
    ws.write_string(3, 0, "Target Bold", bold_i2)
    ws.write_number(3, 1, 1)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("left_edge_then_up_scan",))
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!B4",
            sheet="S",
            row=4,
            col=2,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["Target Bold"]
        assert col == []
    finally:
        wv.close()


def test_left_edge_scan_resets_labels_when_crossing_non_year_numbers(tmp_path) -> None:
    """Crossing a non-year number resets collected labels and keeps scanning."""
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
        assert row == ["A-label"]
        assert col == []
    finally:
        wv.close()


def test_top_edge_scan_resets_labels_when_crossing_non_year_numbers(tmp_path) -> None:
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
        assert col == ["top"]
    finally:
        wv.close()


def test_full_row_scan_collects_text_across_intervening_numbers(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "full_row_scan.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "A-label")
    ws.write_number(0, 1, 999)
    ws.write_string(0, 2, "C-label")
    ws.write_number(0, 3, 1)
    ws.write_string(0, 4, "E-label")
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("full_row_scan",))
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!D1",
            sheet="S",
            row=1,
            col=4,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["E-label", "C-label", "A-label"]
        assert col == []
    finally:
        wv.close()


def test_full_column_scan_collects_text_across_intervening_numbers(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "full_column_scan.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "top")
    ws.write_number(1, 0, 999)
    ws.write_string(2, 0, "mid")
    ws.write_number(3, 0, 1)
    ws.write_string(4, 0, "bottom")
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("full_column_scan",))
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!A4",
            sheet="S",
            row=4,
            col=1,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == []
        assert col == ["bottom", "mid", "top"]
    finally:
        wv.close()


def test_merge_policy_append_dedupe_reverse_reverses_row_labels(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "merge_policy_reverse_row.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "A-label")
    ws.write_number(0, 1, 999)
    ws.write_string(0, 2, "C-label")
    ws.write_number(0, 3, 1)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(
            enabled=True,
            merge_policy="append_dedupe_reverse",
            fallback_behaviors=("full_row_scan", "left_edge_scan"),
        )
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!D1",
            sheet="S",
            row=1,
            col=4,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["A-label", "C-label"]
        assert col == []
    finally:
        wv.close()


def test_merge_policy_append_dedupe_reverse_reverses_column_labels(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "merge_policy_reverse_col.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_string(0, 0, "top")
    ws.write_number(1, 0, 999)
    ws.write_string(2, 0, "mid")
    ws.write_number(3, 0, 1)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(
            enabled=True,
            merge_policy="append_dedupe_reverse",
            fallback_behaviors=("full_column_scan", "top_edge_scan"),
        )
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!A4",
            sheet="S",
            row=4,
            col=1,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == []
        assert col == ["top", "mid"]
    finally:
        wv.close()


def test_right_edge_scan_collects_text_to_the_right(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "right_edge_scan.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_number("A1", 200)
    ws.write_string("B1", "million")
    ws.write_string("C1", "dollars")
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("right_edge_scan",))
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!A1",
            sheet="S",
            row=1,
            col=1,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["million", "dollars"]
        assert col == []
    finally:
        wv.close()


def test_bottom_edge_scan_collects_text_below(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "bottom_edge_scan.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_number("A1", 200)
    ws.write_string("A2", "Source: CIA Factbook, 2012")
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("bottom_edge_scan",))
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!A1",
            sheet="S",
            row=1,
            col=1,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == ["Source: CIA Factbook, 2012"]
        assert col == []
    finally:
        wv.close()


def test_top_edge_then_left_scan_collects_column_hierarchy(tmp_path) -> None:
    pytest.importorskip("xlsxwriter")
    import xlsxwriter

    path = tmp_path / "top_edge_then_left_scan.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    i1 = wb.add_format({"indent": 1})
    i2 = wb.add_format({"indent": 2})
    ws.write_string(0, 0, "Country")
    ws.write_string(0, 1, "GDP", i2)
    ws.write_string(0, 2, "Real", i1)
    ws.write_number(1, 2, 1)
    wb.close()

    import fastpyxl

    wv = fastpyxl.load_workbook(path, data_only=True)
    try:
        wsv = wv["S"]
        cfg = LabelDetectionConfig(enabled=True, fallback_behaviors=("top_edge_then_left_scan",))
        reg = build_label_behavior_registry(None)
        st = LabelDetectionState()
        row, col = collect_labels_for_node(
            key="S!C2",
            sheet="S",
            row=2,
            col=3,
            cfg=cfg,
            registry=reg,
            state=st,
            ws_values=wsv,
            ws_formulas=wsv,
        )
        assert row == []
        assert col == ["Country", "GDP", "Real"]
    finally:
        wv.close()
