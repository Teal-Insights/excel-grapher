"""Cross-sheet reference parsing tests."""

from __future__ import annotations

from excel_grapher.grapher.range_compression.ref_parser import (
    AbsRangeRef,
    parse_ref_streams,
    range_ref_to_keys,
)


def test_parse_cross_sheet_cell_refs() -> None:
    refs = parse_ref_streams("=Data!B3*Data!C3", default_sheet="Report")
    assert len(refs) == 2
    assert refs[0].sheet == "Data"
    assert refs[1].sheet == "Data"


def test_parse_cross_sheet_qualified_range() -> None:
    streams = parse_ref_streams("=SUM(Data!E3:Data!$E$11)", default_sheet="Report")
    assert len(streams) == 1
    ref = streams[0]
    assert isinstance(ref, AbsRangeRef)
    keys = range_ref_to_keys(ref, default_sheet="Report")
    assert all(k.startswith("Data!") for k in keys)
    assert "Data!E3" in keys
    assert "Data!E11" in keys


def test_parse_cross_sheet_vlookup_table() -> None:
    streams = parse_ref_streams(
        "=VLOOKUP(J3,Data!$M$3:Data!$N$7,2,FALSE)",
        default_sheet="Report",
    )
    assert len(streams) == 2
    table = streams[1]
    assert isinstance(table, AbsRangeRef)
    assert table.start_abs_col and table.end_abs_col
