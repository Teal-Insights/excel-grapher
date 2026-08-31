"""`create_dependency_graph` traces dependencies from real-style workbooks (integration).

Uses generated and fixture `.xlsx` files to assert formula chains, array formulas,
and target selection produce the graph topology library users depend on.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Literal

import fastpyxl
import pytest
import xlsxwriter
from fastpyxl.worksheet.formula import ArrayFormula

from excel_grapher import DynamicRefConfig, create_dependency_graph
from excel_grapher.core.cell_types import RealBetween


def _fixture_path(name: str) -> Path:
    return Path(__file__).parent / "data" / name


def _create_fixture_workbook(path: Path) -> None:
    """Create a small workbook with a simple dependency chain.

    - Sheet1!A1 = 2           (leaf)
    - Sheet1!A2 = 3           (leaf)
    - Sheet1!A3 = =A1 + A2    (formula depends on A1, A2)
    - Sheet1!A4 = =A3 * 2     (formula depends on A3)
    """
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"].value = 2
    ws["A2"].value = 3
    ws["A3"].value = "=A1+A2"
    ws["A4"].value = "=A3*2"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def test_create_dependency_graph_traces_dependencies(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _create_fixture_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=False)

    assert "Sheet1!A4" in graph
    assert "Sheet1!A3" in graph
    assert "Sheet1!A2" in graph
    assert "Sheet1!A1" in graph

    assert graph.get_dependencies("Sheet1!A4") == {"Sheet1!A3"}
    assert graph.get_dependencies("Sheet1!A3") == {"Sheet1!A1", "Sheet1!A2"}
    assert graph.get_dependencies("Sheet1!A2") == set()
    assert graph.get_dependencies("Sheet1!A1") == set()


def test_evaluation_order_is_dependency_first(tmp_path: Path) -> None:
    excel_path = tmp_path / "simple_chain.xlsx"
    _create_fixture_workbook(excel_path)

    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=False)
    order = graph.evaluation_order()

    assert order.index("Sheet1!A1") < order.index("Sheet1!A3")
    assert order.index("Sheet1!A2") < order.index("Sheet1!A3")
    assert order.index("Sheet1!A3") < order.index("Sheet1!A4")


def test_range_dependencies_are_expanded(tmp_path: Path) -> None:
    """Expand Excel range references to individual cell dependencies.

    Ensures we don't miss intermediate inputs inside SUM/MIN/MAX/etc.
    """
    excel_path = tmp_path / "range_chain.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"].value = 1
    ws["A2"].value = 2
    ws["A3"].value = 3
    ws["A4"].value = "=SUM(A1:A3)"

    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=False)
    assert graph.get_dependencies("Sheet1!A4") == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}


def test_cross_sheet_range_dependencies_are_expanded(tmp_path: Path) -> None:
    excel_path = tmp_path / "cross_sheet_range.xlsx"
    wb = fastpyxl.Workbook()
    s1 = wb.active
    s1.title = "Sheet1"
    s2 = wb.create_sheet("Sheet 2")

    s2["A1"].value = 10
    s2["A2"].value = 20
    s2["B1"].value = 30
    s2["B2"].value = 40

    s1["A1"].value = "x"
    s1["A2"].value = "=SUM('Sheet 2'!A1:B2)"

    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A2"], load_values=False)
    # Sheet names with spaces are quoted in keys to match Excel formula syntax
    assert graph.get_dependencies("Sheet1!A2") == {
        "'Sheet 2'!A1",
        "'Sheet 2'!A2",
        "'Sheet 2'!B1",
        "'Sheet 2'!B2",
    }


def test_named_range_is_resolved(tmp_path: Path) -> None:
    excel_path = tmp_path / "named_range.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["A1"].value = 41
    ws["A2"].value = "=MyInput+1"

    # Define name: MyInput -> Sheet1!$A$1
    from fastpyxl.workbook.defined_name import DefinedName

    wb.defined_names.add(DefinedName("MyInput", attr_text="Sheet1!$A$1"))

    wb.save(excel_path)
    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A2"], load_values=False)
    assert graph.get_dependencies("Sheet1!A2") == {"Sheet1!A1"}


def test_load_values_reads_cached_formula_results(tmp_path: Path) -> None:
    """When load_values=True, formula nodes should include cached computed values.

    We generate the workbook with XlsxWriter so cached results are embedded.
    """
    import xlsxwriter

    excel_path = tmp_path / "cached_values.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")

    # A1=2, A2=3
    ws.write_number(0, 0, 2)
    ws.write_number(1, 0, 3)

    # A3 = A1+A2 (cached result 5)
    ws.write_formula(2, 0, "=A1+A2", None, 5)
    # A4 = A3*2 (cached result 10)
    ws.write_formula(3, 0, "=A3*2", None, 10)

    wb.close()

    graph = create_dependency_graph(excel_path, ["Sheet1!A4"], load_values=True)

    n3 = graph.get_node("Sheet1!A3")
    n4 = graph.get_node("Sheet1!A4")
    assert n3 is not None and n4 is not None
    assert n3.value == 5
    assert n4.value == 10


def test_array_formula_cells_surface_formula_text(tmp_path: Path) -> None:
    excel_path = tmp_path / "array_formula.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["B1"].value = 1
    ws["B2"].value = 2
    ws["B3"].value = 3

    cell = ws["A1"]
    cell.value = ArrayFormula("A1:A1", "=SUM(B1:B3)")

    wb.save(excel_path)
    wb.close()

    wb_formula = fastpyxl.load_workbook(excel_path, data_only=False, read_only=True)
    try:
        raw = wb_formula["Sheet1"]["A1"].value
        assert isinstance(raw, ArrayFormula)
    finally:
        wb_formula.close()

    graph = create_dependency_graph(
        excel_path, ["Sheet1!A1"], load_values=False, store_raw_formula=True
    )
    node = graph.get_node("Sheet1!A1")
    assert node is not None
    assert node.is_leaf is False
    assert node.formula is not None
    assert node.normalized_formula is not None
    assert node.formula.startswith("=")
    assert node.value is None
    assert node.is_array_formula is True
    assert node.array_formula_ref == "A1:A1"


def test_parse_target_handles_quoted_sheet_name(tmp_path: Path) -> None:
    """Parse target strings with quoted sheet names.

    Keys in the graph use quoted format (e.g. `'Sheet Name'!A1`) to match Excel syntax.
    """
    excel_path = tmp_path / "quoted_sheet.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "My Sheet"

    ws["A1"].value = 42

    wb.save(excel_path)
    wb.close()

    # Target uses quoted format as Excel would show it
    graph = create_dependency_graph(excel_path, ["'My Sheet'!A1"], load_values=False)

    # Keys are quoted when sheet names contain spaces
    assert "'My Sheet'!A1" in graph
    node = graph.get_node("'My Sheet'!A1")
    assert node is not None
    assert node.value == 42
    # Node.sheet stores the unquoted name
    assert node.sheet == "My Sheet"


def test_rejects_workbook_instance_input(tmp_path: Path) -> None:
    """Reject pre-loaded `fastpyxl.Workbook` instances.

    The builder loads from a path with `keep_formula_cache` when values are
    requested, so callers must pass a path rather than a pre-loaded workbook.
    """
    excel_path = tmp_path / "instance_input.xlsx"
    _create_fixture_workbook(excel_path)
    wb = fastpyxl.load_workbook(excel_path, data_only=False)

    with pytest.raises(TypeError, match="path"):
        create_dependency_graph(wb, ["Sheet1!A4"], load_values=True)


def test_offset_invalid_base_error_includes_cell_address(tmp_path: Path) -> None:
    """Include the cell address in OFFSET base-reference errors.

    When OFFSET's base argument is not a cell/range reference, the `ValueError`
    should name the sheet-qualified address for easy diagnosis.
    """
    excel_path = tmp_path / "offset_bad_base.xlsx"
    wb = xlsxwriter.Workbook(excel_path)
    ws = wb.add_worksheet("Sheet1")
    # OFFSET(1,0,0) — base is a literal number, not a cell reference
    ws.write_formula(0, 0, "=OFFSET(1,0,0)", None, 0)
    wb.close()

    with pytest.raises(ValueError, match="Sheet1") as exc_info:
        create_dependency_graph(
            excel_path, ["Sheet1!A1"], load_values=False, use_cached_dynamic_refs=True
        )
    assert "A1" in str(exc_info.value)


def test_absolute_cross_sheet_refs_no_spurious_same_sheet_edge_issue_154(
    tmp_path: Path,
) -> None:
    """Regression for gh #154: Sheet!$Col$Row must not add CurrentSheet!ColRow."""
    cases: list[tuple[str, set[str], str]] = [
        ("=Inputs!$B$5", {"Inputs!B5"}, "Inputs"),
        ("=A1+Inputs!$B$5", {"Engine!A1", "Inputs!B5"}, "Inputs"),
        ("=IF(A1>=Inputs!$B$5,1,0)", {"Engine!A1", "Inputs!B5"}, "Inputs"),
        ("=A1+Inputs!B5", {"Engine!A1", "Inputs!B5"}, "Inputs"),
        ("='Input Sheet'!$B$5", {"'Input Sheet'!B5"}, "Input Sheet"),
        ("=A1+'Input Sheet'!$B$5", {"Engine!A1", "'Input Sheet'!B5"}, "Input Sheet"),
    ]
    for idx, (formula, expected, dep_sheet) in enumerate(cases):
        excel_path = tmp_path / f"issue_154_case_{idx}.xlsx"
        wb = fastpyxl.Workbook()
        default = wb.active
        if default is not None:
            wb.remove(default)
        engine = wb.create_sheet("Engine")
        dep_ws = wb.create_sheet(dep_sheet)
        engine["A1"] = 1
        engine["A2"] = formula
        dep_ws["B5"] = 7
        wb.save(excel_path)
        wb.close()

        graph = create_dependency_graph(excel_path, ["Engine!A2"], load_values=False)
        assert graph.get_dependencies("Engine!A2") == expected, (
            f"formula={formula!r} expected={sorted(expected)!r} "
            f"actual={sorted(graph.get_dependencies('Engine!A2'))!r}"
        )


@pytest.mark.parametrize(
    ("current_sheet", "dep_sheet", "formula", "expected"),
    [
        ("S1", "S2", "=S2!B5", {"S2!B5"}),
        ("S1", "S2", "=1+S2!B5", {"S2!B5"}),
        ("S1", "S2", "='S2'!B5", {"S2!B5"}),
        ("S1", "S2", "='S2'!$B$5", {"S2!B5"}),
        ("Sheet1", "AA1", "=AA1!B5", {"AA1!B5"}),
        ("Sheet1", "AA1", "=AA1!$B$5", {"AA1!B5"}),
        ("Sheet1", "Inputs", "=Inputs!B5", {"Inputs!B5"}),
    ],
    ids=[
        "s2-ref",
        "s2-ref-with-prefix",
        "quoted-s2-ref",
        "quoted-s2-absolute-ref",
        "aa1-ref",
        "aa1-absolute-ref",
        "inputs-control",
    ],
)
def test_address_like_sheet_names_do_not_add_same_sheet_alias_edges_issue_155(
    tmp_path: Path,
    current_sheet: str,
    dep_sheet: str,
    formula: str,
    expected: set[str],
) -> None:
    excel_path = tmp_path / "issue_155_case.xlsx"
    wb = fastpyxl.Workbook()
    default = wb.active
    if default is not None:
        wb.remove(default)

    current = wb.create_sheet(current_sheet)
    dep = wb.create_sheet(dep_sheet)
    current["A2"] = formula
    dep["B5"] = 7

    wb.save(excel_path)
    wb.close()

    target = f"{current_sheet}!A2"
    graph = create_dependency_graph(excel_path, [target], load_values=False)
    assert graph.get_dependencies(target) == expected, (
        f"formula={formula!r} expected={sorted(expected)!r} "
        f"actual={sorted(graph.get_dependencies(target))!r}"
    )


def test_cross_sheet_offset_dynamic_dependencies_include_possible_columns_issue_162(
    tmp_path: Path,
) -> None:
    """Regression for gh #162: cross-sheet OFFSET should include all constrained targets."""
    excel_path = tmp_path / "issue_162_cross_sheet_offset.xlsx"
    wb = fastpyxl.Workbook()
    inputs = wb.active
    assert inputs is not None
    inputs.title = "Inputs"
    engine = wb.create_sheet("Engine")

    inputs["B22"] = 1
    inputs["B26"] = 10
    inputs["C26"] = 20
    inputs["D26"] = 30
    engine["B9"] = "=OFFSET(Inputs!$B$26,0,Inputs!$B$22-1)"

    wb.save(excel_path)
    wb.close()

    constraints = {
        "Inputs!B22": Literal[1, 2, 3],
        "Inputs!B26": Annotated[float, RealBetween(-30.0, 30.0)],
        "Inputs!C26": Annotated[float, RealBetween(-30.0, 30.0)],
        "Inputs!D26": Annotated[float, RealBetween(-30.0, 30.0)],
    }

    graph = create_dependency_graph(
        excel_path,
        ["Engine!B9"],
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_constraints(constraints, {}),
    )

    assert set(graph.leaf_keys()) >= {
        "Inputs!B22",
        "Inputs!B26",
        "Inputs!C26",
        "Inputs!D26",
    }
