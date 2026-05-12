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
