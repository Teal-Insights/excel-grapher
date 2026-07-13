"""Same-sheet ranges converge to a single sheet prefix (#376)."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import format_range_key, normalize_key
from excel_grapher.core.formula_ast import RangeNode, parse
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher.parser import FormulaNormalizer


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
