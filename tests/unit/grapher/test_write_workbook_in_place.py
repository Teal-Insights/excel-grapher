from __future__ import annotations

import zipfile
from pathlib import Path

import fastpyxl
import pytest
import xlsxwriter
from fastpyxl.styles import Border, Font, PatternFill, Side
from fastpyxl.workbook.defined_name import DefinedName
from fastpyxl.worksheet.datavalidation import DataValidation
from fastpyxl.worksheet.formula import ArrayFormula

from excel_grapher.core.formula_ast import parse_preserving_axes
from excel_grapher.exporter.projection import IdentityTransitCompression
from excel_grapher.grapher import (
    DependencyGraph,
    build_graph_cache_meta,
    create_dependency_graph,
    dump_graph,
    load_graph,
    save_graph_cache,
    try_load_graph_cache,
    warm_formula_shapes,
    write_workbook,
)
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.series_bindings import validate_bindings_document

_LIC_DSF_XLSM = Path("tests/fixtures/lic_dsf/lic-dsf-template-2025-08-12.xlsm")


def _write_chain_workbook(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["B1"] = "=A1"
    ws["C1"] = "=B1*2"
    ws["D1"] = 5
    ws["Z99"] = "off-closure"
    wb.save(path)
    wb.close()


def test_extracted_keys_recorded_and_preserved_across_edits_and_projection(
    tmp_path: Path,
) -> None:
    source = tmp_path / "chain.xlsx"
    _write_chain_workbook(source)

    assert DependencyGraph().extracted_keys is None

    graph = create_dependency_graph(source, ["Sheet1!C1"])
    assert graph.extracted_keys == frozenset({"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"})

    # Chained moves do not pollute extracted_keys with transient stops.
    graph.move_node("Sheet1!C1", "Sheet1!C2")
    graph.move_node("Sheet1!C2", "Sheet1!C3")
    assert graph.extracted_keys == frozenset({"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"})
    assert graph.copy().extracted_keys == frozenset({"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"})

    # Materializing a previously unvisited workbook cell extends extracted_keys.
    graph.replace_node_formula("Sheet1!C3", "=B1+D1", None, workbook=source)
    assert graph.extracted_keys == frozenset({"Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"})

    projection = IdentityTransitCompression().project(graph)
    assert projection.extracted_keys == frozenset(
        {"Sheet1!A1", "Sheet1!B1", "Sheet1!C1", "Sheet1!D1"}
    )
    assert "Sheet1!B1" not in projection


def test_extracted_keys_round_trip_json_cache_and_pickle(tmp_path: Path) -> None:
    source = tmp_path / "chain.xlsx"
    _write_chain_workbook(source)

    targets = ["Sheet1!C1"]
    graph = create_dependency_graph(source, targets)
    graph.move_node("Sheet1!C1", "Sheet1!C5")

    # JSON cache round-trip (including vacated key C1 absent from current nodes).
    cache_path = tmp_path / "graph.json"
    meta = build_graph_cache_meta(source, targets)
    save_graph_cache(cache_path, graph, meta)
    from_json = try_load_graph_cache(cache_path, expected_meta=meta)
    assert from_json is not None
    assert from_json.extracted_keys == frozenset({"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"})

    # Multipart pickle round-trip (including vacated key C1 absent from current nodes).
    pkl_path = tmp_path / "graph.pkl"
    dump_graph(graph, pkl_path)
    from_pkl = load_graph(pkl_path)
    assert from_pkl.extracted_keys == frozenset({"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"})


def test_strip_validation_coords_splits_rectangles_and_drops_empty_rules() -> None:
    from excel_grapher.grapher.cell_validation import strip_validation_coords

    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws.add_data_validation(
        DataValidation(type="whole", operator="between", formula1="1", formula2="9", sqref="A1:C3")
    )
    ws.add_data_validation(DataValidation(type="list", formula1='"x,y"', sqref="E5"))
    ws.add_data_validation(
        DataValidation(type="decimal", operator="greaterThan", formula1="0", sqref="G1:G4")
    )

    strip_validation_coords({"Sheet1": ws}, {"Sheet1": {"B2", "E5"}})

    dvs = ws.data_validations.dataValidation
    assert len(dvs) == 2

    covered_first: list[tuple[int, int]] = []
    for rng in dvs[0].sqref.ranges:
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                covered_first.append((col, row))
    expected_border = {
        (col, row) for col in range(1, 4) for row in range(1, 4) if (col, row) != (2, 2)
    }
    assert len(covered_first) == 8
    assert set(covered_first) == expected_border
    assert str(dvs[1].sqref) == "G1:G4"


def _write_rich_workbook(path: Path) -> None:
    wb = fastpyxl.Workbook()
    summary = wb.active
    assert summary is not None
    summary.title = "Summary"
    summary["A1"] = "Executive Dashboard"
    summary["B2"] = "=Sheet1!B1+100"

    ws = wb.create_sheet("Sheet1")
    ws["A1"] = 0.10
    ws["A1"].font = Font(bold=True, color="00FF0000")
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="00FFFF00")
    ws["A1"].number_format = "0.00%"
    ws["B1"] = "=A1*2"
    ws["B1"].number_format = "0.00%"
    ws["G1"] = "Merged Banner"
    ws.merge_cells("G1:H1")
    ws["Z99"] = "Untouched Note"

    wb.defined_names.add(DefinedName(name="FixedMargin", attr_text="0.05"))
    wb.defined_names.add(DefinedName(name="LocalBanner", attr_text="Sheet1!$G$1", localSheetId=1))
    wb.save(path)
    wb.close()


def test_in_place_preserves_styles_sheets_off_closure_cells_defined_names_and_sets_full_calc(
    tmp_path: Path,
) -> None:
    path = tmp_path / "rich.xlsx"
    _write_rich_workbook(path)

    graph = create_dependency_graph(path, ["Sheet1!B1"])
    graph.set_node_value("Sheet1!A1", 0.25)
    graph.replace_node_formula("Sheet1!B1", "=A1*4", None)

    write_workbook(graph, path, in_place=True)

    wb = fastpyxl.load_workbook(path)
    try:
        assert wb.sheetnames == ["Summary", "Sheet1"]
        assert wb["Summary"]["A1"].value == "Executive Dashboard"
        assert wb["Summary"]["B2"].value == "=Sheet1!B1+100"

        ws = wb["Sheet1"]
        assert ws["A1"].value == pytest.approx(0.25)
        assert ws["A1"].font.bold is True
        assert ws["A1"].fill.fill_type == "solid"
        assert ws["A1"].number_format == "0.00%"
        assert ws["B1"].value == "=A1*4"
        assert ws["B1"].number_format == "0.00%"
        assert ws["Z99"].value == "Untouched Note"
        assert [str(rng) for rng in ws.merged_cells.ranges] == ["G1:H1"]

        assert wb.defined_names["FixedMargin"].attr_text == "0.05"
        assert ws.defined_names["LocalBanner"].attr_text == "Sheet1!$G$1"
        assert wb.calculation.fullCalcOnLoad is True
    finally:
        wb.close()


def test_template_mode_patches_into_new_destination_and_honors_overwrite(
    tmp_path: Path,
) -> None:
    source = tmp_path / "template.xlsx"
    dest = tmp_path / "patched.xlsx"
    _write_rich_workbook(source)
    original_bytes = source.read_bytes()

    graph = create_dependency_graph(source, ["Sheet1!B1"])
    graph.set_node_value("Sheet1!A1", 0.5)

    write_workbook(graph, dest, template=source)
    assert source.read_bytes() == original_bytes

    wb = fastpyxl.load_workbook(dest)
    try:
        assert wb["Sheet1"]["A1"].value == pytest.approx(0.5)
        assert wb["Sheet1"]["Z99"].value == "Untouched Note"
        assert wb["Summary"]["A1"].value == "Executive Dashboard"
    finally:
        wb.close()

    with pytest.raises(FileExistsError, match="patched.xlsx"):
        write_workbook(graph, dest, template=source)

    graph.set_node_value("Sheet1!A1", 0.75)
    write_workbook(graph, dest, template=source, overwrite=True)
    wb2 = fastpyxl.load_workbook(dest)
    try:
        assert wb2["Sheet1"]["A1"].value == pytest.approx(0.75)
    finally:
        wb2.close()


def test_in_place_path_and_extension_validation(tmp_path: Path) -> None:
    source = tmp_path / "src.xlsx"
    _write_chain_workbook(source)
    graph = create_dependency_graph(source, ["Sheet1!C1"])

    with pytest.raises(FileNotFoundError):
        write_workbook(graph, tmp_path / "missing.xlsx", in_place=True)

    subdir = tmp_path / "dir_dest"
    subdir.mkdir()
    with pytest.raises(IsADirectoryError):
        write_workbook(graph, subdir, in_place=True)

    other = tmp_path / "other.xlsx"
    _write_chain_workbook(other)
    with pytest.raises(ValueError, match="in_place"):
        write_workbook(graph, source, in_place=True, template=other)

    with pytest.raises(ValueError, match=r"\.xlsm"):
        write_workbook(graph, tmp_path / "out.xlsm", template=source)


def test_in_place_preserves_xlsm_vba_project(tmp_path: Path) -> None:
    with zipfile.ZipFile(_LIC_DSF_XLSM) as zf:
        vba_bytes = zf.read("xl/vbaProject.bin")
    vba_bin = tmp_path / "vbaProject.bin"
    vba_bin.write_bytes(vba_bytes)

    xlsm_path = tmp_path / "macro_book.xlsm"
    xwb = xlsxwriter.Workbook(xlsm_path)
    xwb.add_vba_project(str(vba_bin))
    xws = xwb.add_worksheet("Sheet1")
    xws.write_number(0, 0, 7)
    xws.write_formula(0, 1, "=A1*3", None, 21)
    xws.write_string(4, 4, "Keep macro note")
    xwb.close()

    graph = create_dependency_graph(xlsm_path, ["Sheet1!B1"])
    graph.set_node_value("Sheet1!A1", 12)

    with pytest.raises(ValueError, match=r"\.xlsm"):
        write_workbook(graph, tmp_path / "wrong.xlsx", template=xlsm_path)

    write_workbook(graph, xlsm_path, in_place=True)

    with zipfile.ZipFile(xlsm_path) as zf_after:
        assert zf_after.read("xl/vbaProject.bin") == vba_bytes

    wb = fastpyxl.load_workbook(xlsm_path)
    try:
        assert wb["Sheet1"]["A1"].value == 12
        assert wb["Sheet1"]["B1"].value == "=A1*3"
        assert wb["Sheet1"]["E5"].value == "Keep macro note"
    finally:
        wb.close()


def test_in_place_clears_vacated_cells_across_move_compress_projection_and_caches(
    tmp_path: Path,
) -> None:
    # 1. Chained move_node clears original C1 while keeping border style and transient C2 untouched.
    move_book = tmp_path / "move.xlsx"
    _write_chain_workbook(move_book)
    wb_init = fastpyxl.load_workbook(move_book)
    thin = Side(style="thin", color="00000000")
    wb_init["Sheet1"]["C1"].border = Border(bottom=thin)
    wb_init["Sheet1"]["C2"] = "transient bystander"
    wb_init.save(move_book)
    wb_init.close()

    graph = create_dependency_graph(move_book, ["Sheet1!C1"])
    # Move C1 -> C4 -> C5 (C2 stays untouched because we do not land on it).
    graph.move_node("Sheet1!C1", "Sheet1!C4")
    graph.move_node("Sheet1!C4", "Sheet1!C5")
    write_workbook(graph, move_book, in_place=True)

    wb = fastpyxl.load_workbook(move_book)
    try:
        ws = wb["Sheet1"]
        assert ws["C1"].value is None
        assert ws["C1"].border.bottom.style == "thin"
        assert ws["C4"].value is None
        assert ws["C5"].value == "=B1*2"
        assert ws["C2"].value == "transient bystander"
        assert ws["Z99"].value == "off-closure"
    finally:
        wb.close()

    # 2. In-place compress_identity_transits clears collapsed B1.
    comp_book = tmp_path / "compress.xlsx"
    _write_chain_workbook(comp_book)
    comp_graph = create_dependency_graph(
        comp_book, ["Sheet1!C1"], capture_dependency_provenance=True
    )
    comp_graph.compress_identity_transits()
    write_workbook(comp_graph, comp_book, in_place=True)

    wb = fastpyxl.load_workbook(comp_book)
    try:
        ws = wb["Sheet1"]
        assert ws["B1"].value is None
        assert ws["C1"].value == "=A1*2"
        assert ws["Z99"].value == "off-closure"
    finally:
        wb.close()

    # 3. ProjectionResult clears collapsed B1 in template-patched output.
    proj_src = tmp_path / "proj_src.xlsx"
    proj_out = tmp_path / "proj_out.xlsx"
    _write_chain_workbook(proj_src)
    base_graph = create_dependency_graph(
        proj_src, ["Sheet1!C1"], capture_dependency_provenance=True
    )
    projection = IdentityTransitCompression().project(base_graph)
    write_workbook(projection, proj_out, template=proj_src)

    wb = fastpyxl.load_workbook(proj_out)
    try:
        ws = wb["Sheet1"]
        assert ws["B1"].value is None
        assert ws["C1"].value == "=A1*2"
        assert ws["Z99"].value == "off-closure"
    finally:
        wb.close()

    # 4. Restored graph from JSON cache and pickle still clears vacated cell on in-place write.
    for suffix in ("json", "pkl"):
        book = tmp_path / f"cached_{suffix}.xlsx"
        _write_chain_workbook(book)
        g = create_dependency_graph(book, ["Sheet1!C1"])
        g.move_node("Sheet1!C1", "Sheet1!C6")
        if suffix == "json":
            cpath = tmp_path / "g.json"
            meta = build_graph_cache_meta(book, ["Sheet1!C1"])
            save_graph_cache(cpath, g, meta)
            restored = try_load_graph_cache(cpath, expected_meta=meta)
            assert restored is not None
        else:
            ppath = tmp_path / "g.pkl"
            dump_graph(g, ppath)
            restored = load_graph(ppath)
        write_workbook(restored, book, in_place=True)
        wb = fastpyxl.load_workbook(book)
        try:
            assert wb["Sheet1"]["C1"].value is None
            assert wb["Sheet1"]["C6"].value == "=B1*2"
        finally:
            wb.close()


def test_in_place_refuses_overwriting_occupied_off_closure_cells_and_array_blocks(
    tmp_path: Path,
) -> None:
    path = tmp_path / "collision.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 1
    ws["A2"] = 2
    ws["A3"] = 3
    ws["B1"] = "=A1*10"
    ws["Z99"] = "Do not clobber"
    ws["F1"] = ArrayFormula("F1:F3", "=A1:A3*5")
    wb.save(path)
    wb.close()
    before_bytes = path.read_bytes()

    # Collision with occupied off-closure literal cell Z99.
    g1 = create_dependency_graph(path, ["Sheet1!B1"])
    g1.move_node("Sheet1!B1", "Sheet1!Z99")
    with pytest.raises(ValueError, match="Sheet1!Z99"):
        write_workbook(g1, path, in_place=True)
    assert path.read_bytes() == before_bytes

    # Collision with interior cell F2 of an off-closure ArrayFormula F1:F3.
    g2 = create_dependency_graph(path, ["Sheet1!B1"])
    g2.move_node("Sheet1!B1", "Sheet1!F2")
    with pytest.raises(ValueError, match="F1:F3"):
        write_workbook(g2, path, in_place=True)
    assert path.read_bytes() == before_bytes

    # Planned ArrayFormula spilling onto occupied off-closure cell Z99 fails closed.
    g3 = DependencyGraph(sheet_order=["Sheet1"])
    g3.add_node(
        make_cell_node(
            "Sheet1",
            "Z",
            98,
            is_leaf=False,
            formula_ast=parse_preserving_axes("=1", anchor="Sheet1!Z98"),
            is_array_formula=True,
            array_formula_ref="Z98:Z99",
        )
    )
    with pytest.raises(ValueError, match="Z99"):
        write_workbook(g3, path, in_place=True)
    assert path.read_bytes() == before_bytes


def test_in_place_round_trips_dynamic_array_and_clears_stale_spill_literals(
    tmp_path: Path,
) -> None:
    path = tmp_path / "dyn_in_place.xlsx"
    xwb = xlsxwriter.Workbook(path)
    xws = xwb.add_worksheet("Sheet1")
    for row in range(3):
        xws.write_number(row, 0, row + 1)
    xws.write_dynamic_array_formula("E1:E3", "=A1:A3*2", None, 0)
    xws.write_string(0, 7, "Keep H1")
    xwb.close()

    graph = create_dependency_graph(path, ["Sheet1!E1"])
    graph.set_node_value("Sheet1!A1", 10)
    write_workbook(graph, path, in_place=True)

    wb = fastpyxl.load_workbook(path)
    try:
        ws = wb["Sheet1"]
        assert isinstance(ws["E1"].value, ArrayFormula)
        assert ws["E1"].value.ref == "E1:E3"
        assert ws["E2"].value is None
        assert ws["E3"].value is None
        assert ws["H1"].value == "Keep H1"
    finally:
        wb.close()


@pytest.mark.parametrize(
    ("sheet_name", "off_closure_formula"),
    [
        ("Sheet1", "=A2+1"),
        ("Sheet1", "=SUM(A1:A3)"),
        ("Sheet1", "=SUM(Sheet1!A1:A3)"),
        ("Sheet1", "=SUM(Sheet1!A1:Sheet1!A3)"),
        ("My Sheet", "=SUM('My Sheet'!A1:'My Sheet'!A3)"),
        ("Sheet1", "=SUM(A:A)"),
        ("Sheet1", "=SUM(2:2)"),
        ("Sheet1", "=OFFSET(A1,1,0)"),
    ],
)
def test_in_place_refuses_vacating_cell_referenced_by_off_closure_formula(
    tmp_path: Path,
    sheet_name: str,
    off_closure_formula: str,
) -> None:
    path = tmp_path / "dangling.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = sheet_name
    ws["A1"] = 1
    ws["A2"] = 2
    ws["A3"] = 3
    ws["B2"] = f"='{sheet_name}'!A2*10" if " " in sheet_name else "=A2*10"
    ws["D1"] = off_closure_formula
    wb.save(path)
    wb.close()
    before_bytes = path.read_bytes()

    target = f"'{sheet_name}'!B2" if " " in sheet_name else f"{sheet_name}!B2"
    old_leaf = f"'{sheet_name}'!A2" if " " in sheet_name else f"{sheet_name}!A2"
    new_leaf = f"'{sheet_name}'!C5" if " " in sheet_name else f"{sheet_name}!C5"

    graph = create_dependency_graph(path, [target])
    graph.move_node(old_leaf, new_leaf)

    with pytest.raises(ValueError, match="D1"):
        write_workbook(graph, path, in_place=True)
    assert path.read_bytes() == before_bytes


def test_in_place_defined_name_guard_and_retarget_upsert(tmp_path: Path) -> None:
    path = tmp_path / "named.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["B1"] = "=TaxInput*2"
    wb.defined_names.add(DefinedName(name="TaxInput", attr_text="Sheet1!$A$1"))
    wb.defined_names.add(DefinedName(name="OtherConst", attr_text="0.15"))
    wb.save(path)
    wb.close()

    graph = create_dependency_graph(path, ["Sheet1!B1"])
    graph.move_node("Sheet1!A1", "Sheet1!A5")

    # Without updating TaxInput, vacating Sheet1!A1 fails closed.
    with pytest.raises(ValueError, match="TaxInput"):
        write_workbook(graph, path, in_place=True)

    # Retargeting TaxInput in graph.named_ranges upserts TaxInput and preserves OtherConst.
    assert graph.named_ranges is not None
    graph.named_ranges["TaxInput"] = ("Sheet1", "A5")
    write_workbook(graph, path, in_place=True)

    wb_after = fastpyxl.load_workbook(path)
    try:
        assert wb_after["Sheet1"]["A1"].value is None
        assert wb_after["Sheet1"]["A5"].value == 10
        assert wb_after.defined_names["TaxInput"].attr_text == "Sheet1!$A$5"
        assert wb_after.defined_names["OtherConst"].attr_text == "0.15"
    finally:
        wb_after.close()


def test_in_place_repeated_saves_reconcile_validations_and_default_bindings_workbook(
    tmp_path: Path,
) -> None:
    path = tmp_path / "validated.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    ws["A1"] = 5.0
    ws["A2"] = 6.0
    ws["A3"] = 7.0
    ws["B1"] = "Revenue Label"
    ws["C1"] = "=A1*2"
    ws.add_data_validation(
        DataValidation(
            type="decimal",
            operator="between",
            formula1="0",
            formula2="100",
            sqref="A1:A3",
        )
    )
    wb.save(path)
    wb.close()

    bindings = validate_bindings_document(
        {
            "schema_version": "1.21.0",
            "series": [
                {
                    "id": "rate",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!A1",
                    "layout": "scalar",
                    "input": {},
                    "domain": {"real_between": {"min": 1.0, "max": 20.0}},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [
                            {
                                "concept": "METRIC",
                                "role": "key",
                                "scope": "series",
                                "dtype": "string",
                                "bind": {"kind": "cell", "address": "Sheet1!B1", "read": "string"},
                            }
                        ],
                    },
                    "key": ["METRIC"],
                },
                {
                    "id": "out",
                    "sheet": "Sheet1",
                    "data_range": "Sheet1!C1",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_out"}},
                    "structure": {
                        "measure": {
                            "concept": "OBS_VALUE",
                            "dtype": "float",
                            "bind": {"kind": "data_cell", "read": "float"},
                        },
                        "dimensions": [],
                    },
                    "key": [],
                },
            ],
        }
    )

    # Save twice in place without passing bindings_workbook explicitly.
    for next_val in (8.0, 12.0):
        graph = create_dependency_graph(path, ["Sheet1!C1"], load_values=True)
        graph.set_node_value("Sheet1!A1", next_val)
        write_workbook(graph, path, in_place=True, series_bindings=bindings)

    wb_after = fastpyxl.load_workbook(path)
    try:
        ws_after = wb_after["Sheet1"]
        assert ws_after["A1"].value == pytest.approx(12.0)
        assert ws_after["B1"].value == "Revenue Label"
        dvs = ws_after.data_validations.dataValidation
        # Exactly 2 rules: split off-closure A2:A3 rule + single fresh A1 rule (no duplicates).
        assert len(dvs) == 2
        sqrefs = {str(dv.sqref): (dv.formula1, dv.formula2) for dv in dvs}
        assert sqrefs == {
            "A2:A3": ("0", "100"),
            "A1": ("1", "20"),
        }
    finally:
        wb_after.close()


def test_in_place_shared_formulas_coexist_with_off_closure_shared_formulas(
    tmp_path: Path,
) -> None:
    from excel_grapher.grapher.shared_formulas import SharedFormula

    path = tmp_path / "shared_in_place.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    for row in (1, 2, 3):
        ws[f"A{row}"] = row
    ws["B1"] = SharedFormula(0, ref="B1:B3", text="=A1*2")
    ws["B2"] = SharedFormula(0)
    ws["B3"] = SharedFormula(0)
    ws["D1"] = SharedFormula(1, ref="D1:D2", text="=A1+10")
    ws["D2"] = SharedFormula(1)
    wb.save(path)
    wb.close()

    graph = create_dependency_graph(
        path, ["Sheet1!B1", "Sheet1!B2", "Sheet1!B3"], warm_formula_shapes=True
    )
    graph.set_node_value("Sheet1!A1", 100)
    graph.formula_shapes = warm_formula_shapes(graph)
    write_workbook(graph, path, in_place=True, shared_formulas="require")

    wb_after = fastpyxl.load_workbook(path)
    try:
        ws_after = wb_after["Sheet1"]
        assert ws_after["A1"].value == 100
        assert ws_after["B1"].value == "=A1*2"
        assert ws_after["B2"].value == "=A2*2"
        assert ws_after["B3"].value == "=A3*2"
        assert ws_after["D1"].value == "=A1+10"
        assert ws_after["D2"].value == "=A2+10"
    finally:
        wb_after.close()
