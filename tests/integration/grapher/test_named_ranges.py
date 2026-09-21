"""Named ranges resolve to coordinates and participate correctly in dependency graphs (integration).

Covers `build_named_range_map` and `create_dependency_graph` with defined names
so common authoring patterns (simple names, multi-cell ranges, formulas through names)
remain stable for users importing real workbooks.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Literal

import fastpyxl
import pytest
from fastpyxl.workbook.defined_name import DefinedName

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.core.cell_types import RealBetween
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.grapher.resolver import build_named_range_map


def _new_workbook() -> fastpyxl.Workbook:
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    return wb


def test_named_range_map_allows_simple_range(tmp_path: Path) -> None:
    excel_path = tmp_path / "named_ranges_ok.xlsx"
    wb = _new_workbook()
    wb.defined_names.add(DefinedName("Foo", attr_text="Sheet1!$A$1"))
    wb.defined_names.add(DefinedName("Bar", attr_text="Sheet1!$A$1:$B$2"))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert maps.cell_map["Foo"] == ("Sheet1", "A1")
    assert maps.range_map["Bar"] == ("Sheet1", "A1", "B2")


def test_named_range_range_is_expanded(tmp_path: Path) -> None:
    excel_path = tmp_path / "named_ranges_range_dep.xlsx"
    wb = _new_workbook()
    ws = wb["Sheet1"]
    ws["A1"].value = 1
    ws["B1"].value = 2
    ws["C1"].value = "=SUM(Range1)"
    wb.defined_names.add(DefinedName("Range1", attr_text="Sheet1!$A$1:$B$1"))
    wb.save(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!C1"], load_values=False)
    deps = graph.get_dependencies("Sheet1!C1")
    assert deps == {"Sheet1!A1", "Sheet1!B1"}


def test_named_range_map_raises_on_multi_area(tmp_path: Path) -> None:
    excel_path = tmp_path / "named_ranges_multi.xlsx"
    wb = _new_workbook()
    wb.defined_names.add(DefinedName("Multi", attr_text="Sheet1!$A$1,Sheet1!$B$1"))
    ws = wb["Sheet1"]
    ws["C1"].value = "=SUM(Multi)"
    wb.save(excel_path)

    with pytest.raises(ValueError) as exc:
        create_dependency_graph(excel_path, ["Sheet1!C1"], load_values=False)
    assert "Multi" in str(exc.value)


def test_named_range_map_resolves_offset_counta_formula(tmp_path: Path) -> None:
    """Formula-based name OFFSET(Sheet!A1,0,0,COUNTA(Sheet!A:A),COUNTA(Sheet!1:1)) resolves to range."""
    excel_path = tmp_path / "offset_counta_name.xlsx"
    wb = _new_workbook()
    wb.create_sheet("Country_Information")
    ci = wb["Country_Information"]
    ci["A1"].value = 1
    ci["A2"].value = 2
    ci["B1"].value = 3
    ci["C1"].value = 4
    attr = "OFFSET(Country_Information!$A$1,0,0,COUNTA(Country_Information!$A:$A),COUNTA(Country_Information!$1:$1))"
    wb.defined_names.add(DefinedName("DSF__Country_Info", attr_text=attr))
    wb.save(excel_path)

    wb_loaded = fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    maps = build_named_range_map(wb_loaded)
    assert "DSF__Country_Info" in maps.range_map
    sheet, start, end = maps.range_map["DSF__Country_Info"]
    assert sheet == "Country_Information"
    assert start == "A1"
    assert end == "C2"


def test_named_range_map_resolves_offset_counta_fixed_width(tmp_path: Path) -> None:
    """OFFSET with COUNTA height and literal width (e.g. DSF__COMMODITY_TABLE) resolves when result is wider than sheet used range."""
    excel_path = tmp_path / "offset_counta_fixed_width.xlsx"
    wb = _new_workbook()
    wb.create_sheet("COM")
    com = wb["COM"]
    com["A1"].value = 1
    com["A2"].value = 2
    com["A3"].value = "header"
    attr = "OFFSET(COM!$A$3,0,0,COUNTA(COM!$A:$A),7)"
    wb.defined_names.add(DefinedName("DSF__COMMODITY_TABLE", attr_text=attr))
    wb.save(excel_path)

    wb_loaded = fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    maps = build_named_range_map(wb_loaded)
    assert "DSF__COMMODITY_TABLE" in maps.range_map
    sheet, start, end = maps.range_map["DSF__COMMODITY_TABLE"]
    assert sheet == "COM"
    assert start == "A3"
    assert end == "G5"


def test_dependency_graph_expands_formula_based_named_range(tmp_path: Path) -> None:
    """A formula that references an OFFSET/COUNTA defined name resolves without ValueError."""
    excel_path = tmp_path / "graph_offset_name.xlsx"
    wb = _new_workbook()
    wb.create_sheet("Country_Information")
    ci = wb["Country_Information"]
    ci["A1"].value = "a"
    ci["A2"].value = "b"
    ci["B1"].value = "c"
    ci["C1"].value = "d"
    wb.defined_names.add(
        DefinedName(
            "DSF__Country_Info",
            attr_text="OFFSET(Country_Information!$A$1,0,0,COUNTA(Country_Information!$A:$A),COUNTA(Country_Information!$1:$1))",
        )
    )
    ws = wb["Sheet1"]
    ws["D1"].value = "=COUNTA(DSF__Country_Info)"
    wb.save(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!D1"], load_values=False)
    deps = graph.get_dependencies("Sheet1!D1")
    assert "Country_Information!A1" in deps
    assert "Country_Information!B1" in deps
    assert "Country_Information!C1" in deps
    assert "Country_Information!A2" in deps


def test_normalized_formula_resolves_range_named_range(tmp_path: Path) -> None:
    """Normalized formulas should expand range-based named ranges for codegen."""
    excel_path = tmp_path / "named_range_normalized.xlsx"
    wb = _new_workbook()
    ws = wb["Sheet1"]
    ws["A1"].value = 1
    ws["B1"].value = 2
    # Formula that uses a range-based defined name as table_array
    ws["C1"].value = "=VLOOKUP(Sheet1!A1, NumRange, 2, FALSE())"
    wb.defined_names.add(DefinedName("NumRange", attr_text="Sheet1!$A$1:$B$1"))
    wb.save(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!C1"], load_values=False)
    node = graph.get_node("Sheet1!C1")
    assert node is not None
    # Range-based name should be fully expanded in the derived A1 view so that
    # downstream parsers never see a bare identifier like NumRange.
    formula = node.normalized_formula
    assert formula is not None
    compact = formula.replace(" ", "").replace("$", "")
    assert "NumRange" not in compact
    assert "Sheet1!A1:B1" in compact


def test_named_range_map_resolves_offset_counta_plus_literal(tmp_path: Path) -> None:
    """OFFSET height/width with COUNTA(...)+n (LIC-DSF DSF__DATA_* pattern) resolve to padded ranges."""
    excel_path = tmp_path / "offset_counta_plus.xlsx"
    wb = _new_workbook()
    wb.create_sheet("CPIA")
    cpia = wb["CPIA"]
    # Two non-blank rows in col A, three non-blank header cells → COUNTA+5 => h=7, w=8
    cpia["A1"].value = "Database"
    cpia["B1"].value = "Code"
    cpia["C1"].value = "Year"
    cpia["A2"].value = "book.xlsx"
    cpia["B2"].value = 652
    cpia["C2"].value = 3.5
    attr = "OFFSET(CPIA!$A$1,0,0,COUNTA(CPIA!$A:$A)+5,COUNTA(CPIA!$1:$1)+5)"
    wb.defined_names.add(DefinedName("DSF__DATA_CPIA", attr_text=attr))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert "DSF__DATA_CPIA" in maps.range_map
    sheet, start, end = maps.range_map["DSF__DATA_CPIA"]
    assert sheet == "CPIA"
    assert start == "A1"
    # Must not collapse to the 1x1 anchor A1:A1 (issue #410).
    assert end != "A1"
    assert end == "H7"


def test_named_range_map_resolves_offset_counta_minus_literal(tmp_path: Path) -> None:
    """OFFSET height with COUNTA(...)-n resolves without collapsing to the anchor."""
    excel_path = tmp_path / "offset_counta_minus.xlsx"
    wb = _new_workbook()
    wb.create_sheet("lookup")
    lookup = wb["lookup"]
    lookup["AX4"].value = "a"
    lookup["AX5"].value = "b"
    lookup["AX6"].value = "c"
    attr = "OFFSET(lookup!$AX$4,0,0,COUNTA(lookup!$AX:$AX)-1,1)"
    wb.defined_names.add(DefinedName("SHEETS_TO_HIDE", attr_text=attr))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert "SHEETS_TO_HIDE" in maps.range_map
    sheet, start, end = maps.range_map["SHEETS_TO_HIDE"]
    assert sheet == "lookup"
    assert start == "AX4"
    # COUNTA(AX4:AX6)=3 → height 2 → AX4:AX5
    assert end == "AX5"


def test_named_range_map_omits_offset_when_counta_minus_nonpositive(
    tmp_path: Path,
) -> None:
    """COUNTA(...)-n that yields a non-positive height must not register a poison 1x1."""
    excel_path = tmp_path / "offset_counta_nonpositive.xlsx"
    wb = _new_workbook()
    wb.create_sheet("lookup")
    lookup = wb["lookup"]
    lookup["AX4"].value = "only"
    attr = "OFFSET(lookup!$AX$4,0,0,COUNTA(lookup!$AX:$AX)-1,1)"
    wb.defined_names.add(DefinedName("SHEETS_TO_HIDE", attr_text=attr))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert "SHEETS_TO_HIDE" not in maps.range_map
    assert "SHEETS_TO_HIDE" not in maps.cell_map


def test_named_range_map_omits_offset_with_unevaluable_extent(tmp_path: Path) -> None:
    """Explicit height using unsupported COUNTIF must not collapse to the 1x1 anchor."""
    excel_path = tmp_path / "offset_countif_extent.xlsx"
    wb = _new_workbook()
    wb.create_sheet("lookup")
    lookup = wb["lookup"]
    lookup["AY4"].value = "x"
    lookup["AY5"].value = "y"
    attr = 'OFFSET(lookup!$AY$4,0,0,COUNTIF(lookup!$AY:$AY,"?*")-1,1)'
    wb.defined_names.add(DefinedName("COUNTRY_DROP_DOWN", attr_text=attr))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert "COUNTRY_DROP_DOWN" not in maps.range_map
    assert "COUNTRY_DROP_DOWN" not in maps.cell_map


def test_named_range_map_omits_offset_with_non_arithmetic_binary_extent(
    tmp_path: Path,
) -> None:
    """Concat/comparison binary ops in extents must omit the name, not raise."""
    excel_path = tmp_path / "offset_concat_extent.xlsx"
    wb = _new_workbook()
    wb.create_sheet("S")
    wb["S"]["A1"].value = 1
    wb.defined_names.add(DefinedName("BadConcat", attr_text="OFFSET(S!$A$1,0,0,1&2,1)"))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert "BadConcat" not in maps.range_map
    assert "BadConcat" not in maps.cell_map


def test_named_range_map_resolves_whole_row_defined_names(tmp_path: Path) -> None:
    """Whole-row names expand the implicit columns against the sheet used range."""
    excel_path = tmp_path / "whole_row_names.xlsx"
    wb = _new_workbook()
    ws = wb.active
    ws.title = "Macrofw"
    ws["A5"] = "code"
    ws["B5"] = -1
    ws["A6"] = "FASAOS"
    ws["B6"] = 42
    wb.defined_names.add(DefinedName("MacrofwData", attr_text="Macrofw!$5:$10"))
    wb.defined_names.add(DefinedName("MacrofwDates", attr_text="Macrofw!$5:$5"))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert maps.range_map["MacrofwData"] == ("Macrofw", "A5", "B10")
    assert maps.range_map["MacrofwDates"] == ("Macrofw", "A5", "B5")


def test_named_range_map_resolves_whole_column_defined_names(tmp_path: Path) -> None:
    """Whole-column names expand the implicit rows against the sheet used range."""
    excel_path = tmp_path / "whole_column_names.xlsx"
    wb = _new_workbook()
    ws = wb["Sheet1"]
    ws["A1"] = "hdr"
    ws["B2"] = 1
    ws["C3"] = 2
    wb.defined_names.add(DefinedName("Cols", attr_text="Sheet1!$B:$C"))
    wb.defined_names.add(DefinedName("OneCol", attr_text="Sheet1!$A:$A"))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert maps.range_map["Cols"] == ("Sheet1", "B1", "C3")
    assert maps.range_map["OneCol"] == ("Sheet1", "A1", "A3")


def test_named_range_map_resolves_quoted_and_reversed_whole_axis_names(
    tmp_path: Path,
) -> None:
    """Quoted sheets and reversed endpoints still become a used-range rectangle."""
    excel_path = tmp_path / "quoted_whole_axis.xlsx"
    wb = _new_workbook()
    ws = wb.active
    ws.title = "My Sheet"
    ws["A5"] = "a"
    ws["B6"] = "b"
    wb.defined_names.add(DefinedName("Rows", attr_text="'My Sheet'!$6:$5"))
    wb.defined_names.add(DefinedName("Cols", attr_text="'My Sheet'!$C:$A"))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert maps.range_map["Rows"] == ("My Sheet", "A5", "B6")
    assert maps.range_map["Cols"] == ("My Sheet", "A1", "C6")


def test_named_range_map_omits_whole_axis_name_on_missing_sheet(tmp_path: Path) -> None:
    """Unknown-sheet whole-row/column names stay omitted instead of expanding to XFD."""
    excel_path = tmp_path / "missing_sheet_whole_row.xlsx"
    wb = _new_workbook()
    wb["Sheet1"]["A1"] = 1
    wb.defined_names.add(DefinedName("GhostRows", attr_text="Missing!$5:$10"))
    wb.defined_names.add(DefinedName("GhostCols", attr_text="Missing!$A:$C"))
    wb.save(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert "GhostRows" not in maps.range_map
    assert "GhostCols" not in maps.range_map


def test_create_dependency_graph_resolves_whole_row_vlookup(tmp_path: Path) -> None:
    """VLOOKUP/MATCH over whole-row names extract and evaluate like rectangular names."""
    excel_path = tmp_path / "mcve-whole-row-defined-name.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Macrofw"
    ws["A5"] = "code"
    ws["B5"] = -1
    ws["A6"] = "FASAOS"
    ws["B6"] = 42
    qs = wb.create_sheet("Questions")
    qs["C90"] = '=VLOOKUP("FASAOS",MacrofwData,MATCH(-1,MacrofwDates,0),FALSE)'
    wb.defined_names.add(DefinedName("MacrofwData", attr_text="Macrofw!$5:$10"))
    wb.defined_names.add(DefinedName("MacrofwDates", attr_text="Macrofw!$5:$5"))
    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(
        excel_path,
        targets=["Questions!C90"],
        max_depth=10,
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    node = graph.get_node("Questions!C90")
    assert node is not None
    assert node.normalized_formula is not None
    assert "Macrofw!A5:B10" in node.normalized_formula
    assert "Macrofw!A5:B5" in node.normalized_formula
    deps = graph.get_dependencies("Questions!C90")
    assert "Macrofw!A6" in deps
    assert "Macrofw!B5" in deps
    with FormulaEvaluator(graph) as ev:
        assert ev.evaluate("Questions!C90") == 42


def test_dependency_graph_expands_offset_counta_plus_named_range(tmp_path: Path) -> None:
    """INDEX over a COUNTA+n named range expands to the padded table, not A1:A1."""
    excel_path = tmp_path / "graph_offset_counta_plus.xlsx"
    wb = _new_workbook()
    wb.create_sheet("CPIA")
    cpia = wb["CPIA"]
    cpia["A1"].value = "hdr"
    cpia["B1"].value = "code"
    cpia["C1"].value = "year"
    cpia["A2"].value = "row"
    cpia["B2"].value = 652
    cpia["C2"].value = 3.5
    wb.defined_names.add(
        DefinedName(
            "DSF__DATA_CPIA",
            attr_text="OFFSET(CPIA!$A$1,0,0,COUNTA(CPIA!$A:$A)+5,COUNTA(CPIA!$1:$1)+5)",
        )
    )
    ws = wb["Sheet1"]
    ws["D1"].value = "=INDEX(DSF__DATA_CPIA,2,3)"
    wb.save(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!D1"], load_values=False)
    node = graph.get_node("Sheet1!D1")
    assert node is not None
    assert node.normalized_formula is not None
    assert "CPIA!A1:A1" not in node.normalized_formula
    assert "CPIA!A1:H7" in node.normalized_formula
    deps = graph.get_dependencies("Sheet1!D1")
    assert "CPIA!C2" in deps


def _index_offset_named_start_workbook(path: Path) -> None:
    """INDEX(PREV_DSA, MATCH(...)) where PREV_DSA is OFFSET of named START (#957)."""
    wb = fastpyxl.Workbook()
    data = wb.active
    data.title = "data all"
    data["A1"] = "code"
    data["B1"] = "value"
    data["A2"] = "NFIG_R_GDP"
    data["B2"] = 1.71
    wb.defined_names.add(DefinedName(name="START", attr_text="'data all'!$A$1"))
    wb.defined_names.add(
        DefinedName(
            name="PREV_DSA",
            attr_text="OFFSET(START,0,0,COUNTA('data all'!$A:$A),2)",
        )
    )
    out = wb.create_sheet("Out")
    out["C1"] = "NFIG_R_GDP"
    out["A1"] = "=INDEX(PREV_DSA, MATCH(C1, INDEX(PREV_DSA,,1), 0), 2)"
    wb.save(path)
    wb.close()


def test_named_range_map_resolves_offset_of_named_start_cell(tmp_path: Path) -> None:
    """OFFSET(START, ..., COUNTA, width) resolves when START is another defined name."""
    excel_path = tmp_path / "index_offset_named_start.xlsx"
    _index_offset_named_start_workbook(excel_path)

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert maps.cell_map["START"] == ("data all", "A1")
    assert maps.range_map["PREV_DSA"] == ("data all", "A1", "B2")


def test_named_range_map_resolves_offset_named_start_counta_minus(tmp_path: Path) -> None:
    """OFFSET(START, ..., COUNTA(col)-1, COUNTA(row)) matches the LIC-DSF PREV_DSA shape."""
    excel_path = tmp_path / "index_offset_named_start_counta_minus.xlsx"
    wb = fastpyxl.Workbook()
    data = wb.active
    data.title = "data all"
    data["A1"] = "code"
    data["B1"] = "value"
    data["A2"] = "NFIG_R_GDP"
    data["B2"] = 1.71
    data["A3"] = "OTHER"
    data["B3"] = 0.5
    wb.defined_names.add(DefinedName(name="START", attr_text="'data all'!$A$1"))
    wb.defined_names.add(
        DefinedName(
            name="PREV_DSA",
            attr_text="OFFSET(START,0,0,COUNTA('data all'!$A:$A)-1,COUNTA('data all'!$2:$2))",
        )
    )
    wb.save(excel_path)
    wb.close()

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert maps.cell_map["START"] == ("data all", "A1")
    # COUNTA(A:A)=3 → height 2; COUNTA(row 2)=2 → A1:B2
    assert maps.range_map["PREV_DSA"] == ("data all", "A1", "B2")


def test_named_range_map_resolves_offset_of_offset_named_start(tmp_path: Path) -> None:
    """OFFSET of a formula name that is itself OFFSET of a named start cell."""
    excel_path = tmp_path / "offset_of_offset_named_start.xlsx"
    wb = fastpyxl.Workbook()
    data = wb.active
    data.title = "data all"
    data["A1"] = "code"
    data["B1"] = "value"
    data["A2"] = "NFIG_R_GDP"
    data["B2"] = 1.71
    wb.defined_names.add(DefinedName(name="START", attr_text="'data all'!$A$1"))
    wb.defined_names.add(DefinedName(name="ANCHOR", attr_text="OFFSET(START,0,0,1,1)"))
    wb.defined_names.add(
        DefinedName(
            name="PREV_DSA",
            attr_text="OFFSET(ANCHOR,0,0,COUNTA('data all'!$A:$A),2)",
        )
    )
    wb.save(excel_path)
    wb.close()

    maps = build_named_range_map(
        fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    )
    assert maps.range_map["ANCHOR"] == ("data all", "A1", "A1")
    assert maps.range_map["PREV_DSA"] == ("data all", "A1", "B2")


def test_index_match_over_offset_named_start_extracts(tmp_path: Path) -> None:
    """INDEX of an OFFSET-defined name extracts once MATCH leaves are constrained."""
    excel_path = tmp_path / "index_offset_named_start_extract.xlsx"
    _index_offset_named_start_workbook(excel_path)

    graph = create_dependency_graph(
        excel_path,
        targets=["Out!A1"],
        dynamic_refs=DynamicRefConfig.from_constraints(
            {
                "Out!C1": Literal["NFIG_R_GDP"],
                "'data all'!A1": Literal["code"],
                "'data all'!A2": Literal["NFIG_R_GDP"],
                "'data all'!B1": Literal["value"],
                "'data all'!B2": Annotated[float, RealBetween(0, 10)],
            }
        ),
    )
    assert graph.named_range_ranges is not None
    assert graph.named_range_ranges["PREV_DSA"] == ("data all", "A1", "B2")
    node = graph.get_node("Out!A1")
    assert node is not None
    assert node.normalized_formula is not None
    formula = node.normalized_formula.replace("$", "")
    assert "PREV_DSA" not in formula
    assert "'data all'!A1:B2" in formula
    deps = graph.get_dependencies("Out!A1")
    assert "Out!C1" in deps
    assert "'data all'!B2" in deps
