"""Same-sheet ranges converge to a single sheet prefix (#376)."""

from __future__ import annotations

from pathlib import Path

import fastpyxl
import pytest

from excel_grapher.core.address_keys import format_range_key, normalize_key
from excel_grapher.core.formula_ast import RangeNode, parse
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher.builder import _format_missing_leaves, create_dependency_graph
from excel_grapher.grapher.parser import (
    CellRef,
    FormulaNormalizer,
    parse_cell_refs,
    parse_standalone_cell_refs,
)


@pytest.mark.parametrize(
    ("sheet", "start", "end", "expected"),
    [
        ("Sheet1", "A1", "A3", "Sheet1!A1:A3"),
        ("Sheet1", "A1", "D1", "Sheet1!A1:D1"),
        ("Sheet1", "A1", "B2", "Sheet1!A1:B2"),
        ("My Sheet", "A1", "B2", "'My Sheet'!A1:B2"),
        ("O'Neil", "C4", "D4", "'O''Neil'!C4:D4"),
    ],
)
def test_format_range_key_uses_single_prefix(
    sheet: str, start: str, end: str, expected: str
) -> None:
    assert format_range_key(sheet, start, end) == expected


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("Sheet1!A1:A3", "Sheet1!A1:A3"),
        ("Sheet1!A1:Sheet1!A3", "Sheet1!A1:A3"),
        ("Sheet1!A1:D1", "Sheet1!A1:D1"),
        ("Sheet1!A1:Sheet1!D1", "Sheet1!A1:D1"),
        ("Sheet1!A1:B2", "Sheet1!A1:B2"),
        ("Sheet1!A1:Sheet1!B2", "Sheet1!A1:B2"),
        ("'My Sheet'!A1:B2", "'My Sheet'!A1:B2"),
        ("'My Sheet'!A1:'My Sheet'!B2", "'My Sheet'!A1:B2"),
        ("'Sheet1'!A1:'Sheet1'!A3", "Sheet1!A1:A3"),
        # Cross-sheet stays both-end (both endpoints require a sheet).
        ("Sheet1!A1:Sheet2!D1", "Sheet1!A1:Sheet2!D1"),
    ],
)
def test_normalize_key_collapses_same_sheet_range_to_single_prefix(raw: str, expected: str) -> None:
    assert normalize_key(raw) == expected


class TestFormulaNormalizerSinglePrefix:
    """FormulaNormalizer emits single-prefix same-sheet ranges."""

    def test_local_range(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=SUM(A1:A3)", "Sheet1") == "=SUM(Sheet1!A1:A3)"

    def test_single_prefix_input(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=SUM(Sheet1!A1:A3)", "Sheet1") == "=SUM(Sheet1!A1:A3)"

    def test_both_end_input_collapsed(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=SUM(Sheet1!A1:Sheet1!A3)", "Sheet1") == "=SUM(Sheet1!A1:A3)"

    def test_quoted_sheet_local_and_both_end(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=SUM(A1:B2)", "My Sheet") == "=SUM('My Sheet'!A1:B2)"
        assert n.normalize("=SUM('My Sheet'!A1:'My Sheet'!B2)", "Other") == "=SUM('My Sheet'!A1:B2)"

    def test_cross_sheet_range_keeps_both_endpoints(self) -> None:
        n = FormulaNormalizer()
        assert n.normalize("=SUM(Sheet1!A1:Sheet2!D1)", "Sheet1") == "=SUM(Sheet1!A1:Sheet2!D1)"

    def test_named_range_range_uses_single_prefix(self) -> None:
        n = FormulaNormalizer(named_range_ranges={"MyTable": ("Sheet1", "A1", "B3")})
        assert n.normalize("=SUM(MyTable)", "Sheet1") == "=SUM(Sheet1!A1:B3)"

    def test_multi_row_and_one_row_agree_with_normalize_key(self) -> None:
        n = FormulaNormalizer()
        for formula, sheet in (
            ("=SUM(A1:A3)", "Sheet1"),
            ("=SUM(A1:D1)", "Sheet1"),
            ("=SUM(A1:B2)", "Sheet1"),
            ("=SUM(Sheet1!A1:Sheet1!B2)", "Sheet1"),
        ):
            normalized = n.normalize(formula, sheet)
            # Extract the range argument inside SUM(...).
            inner = normalized.removeprefix("=SUM(").removesuffix(")")
            assert normalize_key(inner) == inner


def test_ast_and_codegen_emit_single_prefix_xl_range() -> None:
    """AST accepts both-end and single-prefix; codegen emits single-prefix."""
    assert parse("=Sheet1!A1:A3") == RangeNode("Sheet1!A1", "Sheet1!A3")
    assert parse("=Sheet1!A1:Sheet1!A3") == RangeNode("Sheet1!A1", "Sheet1!A3")

    gen = CodeGenerator(None)  # type: ignore
    assert gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!A3")) == "xl_range(ctx, 'Sheet1!A1:A3')"
    assert gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!B2")) == "xl_range(ctx, 'Sheet1!A1:B2')"
    assert (
        gen._emit_ast(RangeNode("'My Sheet'!A1", "'My Sheet'!C1"))
        == "xl_range(ctx, \"'My Sheet'!A1:C1\")"
    )


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(A1:A3,B5)",
        "=SUM(Sheet1!A1:A3,B5)",
        "=SUM(Sheet1!A1:Sheet1!A3,B5)",
        "=SUM('My Sheet'!A1:'My Sheet'!B2,C1)",
    ],
)
def test_parse_standalone_cell_refs_masks_range_spans(formula: str) -> None:
    """Bare range endpoints must not appear as standalone local cells."""
    normalized = FormulaNormalizer().normalize(formula, "Sheet1")
    if formula == "=SUM('My Sheet'!A1:'My Sheet'!B2,C1)":
        normalized = FormulaNormalizer().normalize(formula, "Other")

    standalone = parse_standalone_cell_refs(normalized)
    assert all(
        ref.sheet is not None or ref.column + str(ref.row) not in {"A3", "B2"} for ref in standalone
    )
    # Extra non-range cell still present.
    assert any(
        (ref.sheet is None and f"{ref.column}{ref.row}" in {"B5", "C1"})
        or (ref.sheet is not None and f"{ref.column}{ref.row}" in {"B5", "C1"})
        for ref in standalone
    )


def test_local_cell_re_does_not_match_bare_range_endpoint() -> None:
    """Defense in depth: `:` blocks local-cell matching of range ends."""
    refs = parse_cell_refs("=SUM(Sheet1!A1:A3)", allow_unmasked_ranges=True)
    assert CellRef(sheet=None, column="A", row=3) not in refs


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(Sheet1!A1:A3)",
        "=SUM(Sheet1!A1:Sheet1!A3)",
        "=SUM(A1:A3)+B5",
        "=SUM(Sheet1!A:A)",
    ],
)
def test_parse_cell_refs_refuses_unmasked_ranges(formula: str) -> None:
    with pytest.raises(ValueError, match="unmasked range|allow_unmasked_ranges"):
        parse_cell_refs(formula)


def test_parse_cell_refs_allows_unmasked_ranges_opt_in() -> None:
    refs = parse_cell_refs("=SUM(Sheet1!A1:A3)+B5", allow_unmasked_ranges=True)
    assert CellRef(sheet="Sheet1", column="A", row=1) in refs
    assert CellRef(sheet=None, column="B", row=5) in refs


def test_parse_cell_refs_accepts_masked_or_range_free_text() -> None:
    assert parse_cell_refs("=Sheet1!A1+B5") == [
        CellRef(sheet="Sheet1", column="A", row=1),
        CellRef(sheet=None, column="B", row=5),
    ]
    assert parse_standalone_cell_refs("=SUM(Sheet1!A1:A3)+B5") == [
        CellRef(sheet=None, column="B", row=5)
    ]


def test_format_missing_leaves_uses_single_prefix() -> None:
    assert _format_missing_leaves({"Sheet1!C4", "Sheet1!C5", "Sheet1!C6"}) == ["Sheet1!C4:C6"]
    assert _format_missing_leaves({"S!AA100", "S!AB100", "S!AC100"}) == ["S!AA100:AC100"]
    assert _format_missing_leaves({"S!AA10", "S!AA11", "S!AB10", "S!AB11"}) == ["S!AA10:AB11"]
    assert _format_missing_leaves({"My Sheet!A1", "My Sheet!A2"}) == ["'My Sheet'!A1:A2"]


def _write_cross_sheet_sum(path: Path, *, formula: str) -> None:
    wb = fastpyxl.Workbook()
    s1 = wb.active
    assert s1 is not None
    s1.title = "Sheet1"
    s2 = wb.create_sheet("Sheet2")
    s1["A1"].value = 1
    s1["A2"].value = 2
    s1["A3"].value = 3
    s2["B1"].value = formula
    s2["B2"].value = 9
    wb.save(path)
    wb.close()


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(Sheet1!A1:A3)+B2",
        "=SUM(Sheet1!A1:Sheet1!A3)+B2",
        "=SUM(A1:A3)+Sheet2!B2",  # written on Sheet1 in helper below
    ],
)
def test_dep_extraction_single_prefix_matrix_no_phantom_sheet(tmp_path: Path, formula: str) -> None:
    """Dep extraction expands interiors and never mis-sheets bare endpoints."""
    path = tmp_path / "deps.xlsx"
    if formula.startswith("=SUM(A1:A3)"):
        wb = fastpyxl.Workbook()
        s1 = wb.active
        assert s1 is not None
        s1.title = "Sheet1"
        s2 = wb.create_sheet("Sheet2")
        s1["A1"].value = 1
        s1["A2"].value = 2
        s1["A3"].value = 3
        s2["B2"].value = 9
        s1["B1"].value = formula
        wb.save(path)
        wb.close()
        target = "Sheet1!B1"
        expected = {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3", "Sheet2!B2"}
    else:
        _write_cross_sheet_sum(path, formula=formula)
        target = "Sheet2!B1"
        expected = {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3", "Sheet2!B2"}

    graph = create_dependency_graph(path, [target], load_values=False)
    assert graph.get_dependencies(target) == expected
    assert "Sheet2!A3" not in graph.get_dependencies(target)


def test_expand_ranges_false_does_not_mis_sheet_bare_endpoint(tmp_path: Path) -> None:
    """With expand_ranges=False, bare range ends must not become local-sheet deps."""
    path = tmp_path / "no_expand.xlsx"
    _write_cross_sheet_sum(path, formula="=SUM(Sheet1!A1:A3)+B2")

    graph = create_dependency_graph(path, ["Sheet2!B1"], load_values=False, expand_ranges=False)
    deps = graph.get_dependencies("Sheet2!B1")
    assert deps == {"Sheet2!B2"}
    assert "Sheet2!A3" not in deps
    assert "Sheet1!A1" not in deps
    assert "Sheet1!A2" not in deps
    assert "Sheet1!A3" not in deps
