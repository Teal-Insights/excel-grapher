"""Tests for :mod:`excel_grapher.core.formula_normalization` (canonical normalization)."""

from __future__ import annotations

from excel_grapher.core.formula_normalization import (
    normalize_excel_formula,
    prepare_formula,
)
from excel_grapher.grapher.parser import FormulaNormalizer, normalize_formula


def test_normalize_excel_formula_matches_grapher_normalize_formula() -> None:
    f = "=Inputs!C16+CHOOSE(Inputs!$B$22,$B$9,0,0)*C10"
    sheet = "Engine"
    assert normalize_excel_formula(f, sheet) == normalize_formula(f, sheet)


def test_prepare_formula_wraps_normalize_excel_formula() -> None:
    p = prepare_formula("=A1", "S1")
    assert p.normalized_formula == "=S1!A1"


def test_formula_normalizer_uses_same_pipeline_as_core() -> None:
    n = FormulaNormalizer()
    f = "=SUM(A1:B2)+'Other'!C3"
    assert n.normalize(f, "Here") == normalize_excel_formula(f, "Here")


def test_normalize_excel_formula_preserves_cell_like_text_in_string_literals() -> None:
    """Bare refs inside quoted strings must not be sheet-qualified."""
    cases = [
        ('=IF(B5>0,"C13 is large",D5)', '=IF(Sheet1!B5>0,"C13 is large",Sheet1!D5)'),
        ('="value at C14 is "&B6', '="value at C14 is "&Sheet1!B6'),
        ('=B7&" says ""C15"""', '=Sheet1!B7&" says ""C15"""'),
        ('="C17 has data "&B9', '="C17 has data "&Sheet1!B9'),
    ]
    for formula, expected in cases:
        assert normalize_excel_formula(formula, "Sheet1") == expected
