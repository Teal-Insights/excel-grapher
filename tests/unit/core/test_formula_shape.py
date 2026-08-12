"""Tests for parameterized formula AST shape fingerprinting."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.core.formula_ast import (
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    WholeColumnNode,
    WholeRowNode,
    parse,
)
from excel_grapher.core.formula_shape import (
    FormulaShapeSummary,
    fingerprint_formula_shape,
    specialize_formula_shape,
    summarize_formula_shapes,
    summarize_normalized_formulas,
)

_TACO = Path("examples/micro_workbooks/taco_patterns.xlsx")
_FFV2 = Path("examples/micro_workbooks/ffv2.xlsx")
_TACO_TARGETS = (
    "Patterns!D3:D7",
    "Patterns!F3:F7",
    "Patterns!H3:H7",
    "Patterns!K3:K7",
    "Patterns!P3:P7",
)


def test_fingerprint_punches_cell_refs_and_keeps_literals() -> None:
    shape = fingerprint_formula_shape("=Sheet1!B3+0.1*Sheet1!B5+6*Sheet1!B7")
    assert shape.params == (
        CellRefNode("Sheet1!B3"),
        CellRefNode("Sheet1!B5"),
        CellRefNode("Sheet1!B7"),
    )
    other = fingerprint_formula_shape("=Sheet1!C3+0.1*Sheet1!C5+6*Sheet1!C7")
    assert shape.shape_key == other.shape_key
    assert other.params == (
        CellRefNode("Sheet1!C3"),
        CellRefNode("Sheet1!C5"),
        CellRefNode("Sheet1!C7"),
    )


def test_fingerprint_accepts_parsed_ast() -> None:
    ast = parse("=Sheet1!A1+Sheet1!B2")
    shape = fingerprint_formula_shape(ast)
    assert shape.params == (CellRefNode("Sheet1!A1"), CellRefNode("Sheet1!B2"))
    assert "$CELL" in shape.shape_key
    assert "Sheet1!A1" not in shape.shape_key


def test_fingerprint_range_and_function_name_are_part_of_shape() -> None:
    sum_shape = fingerprint_formula_shape("=SUM(Sheet1!E2:E10)")
    avg_shape = fingerprint_formula_shape("=AVERAGE(Sheet1!E2:E10)")
    assert sum_shape.params == (RangeNode("Sheet1!E2", "Sheet1!E10"),)
    assert avg_shape.params == (RangeNode("Sheet1!E2", "Sheet1!E10"),)
    assert sum_shape.shape_key != avg_shape.shape_key
    assert "$RANGE" in sum_shape.shape_key


def test_fingerprint_whole_column_and_row_kinds_differ() -> None:
    col = fingerprint_formula_shape('=MATCH("x",Data!A:A,0)')
    row_ast = FunctionCallNode(
        "MATCH",
        [StringNode("x"), WholeRowNode(sheet="Data", row=1), NumberNode(0.0)],
    )
    row = fingerprint_formula_shape(row_ast)
    assert col.params == (WholeColumnNode(sheet="Data", column="A"),)
    assert row.params == (WholeRowNode(sheet="Data", row=1),)
    assert col.shape_key != row.shape_key
    assert "$WHOLE_COL" in col.shape_key
    assert "$WHOLE_ROW" in row.shape_key


def test_different_literals_yield_different_shapes() -> None:
    a = fingerprint_formula_shape("=Patterns!P3+1")
    b = fingerprint_formula_shape("=Patterns!P3+2")
    assert a.shape_key != b.shape_key


def test_specialize_round_trips_original_ast() -> None:
    formula = "=Sheet1!B3+0.1*Sheet1!B5+6*Sheet1!B7"
    original = parse(formula)
    shape = fingerprint_formula_shape(original)
    rebuilt = specialize_formula_shape(shape.skeleton, shape.params)
    assert rebuilt == original


def test_specialize_rebounds_params_onto_shared_skeleton() -> None:
    shape_b = fingerprint_formula_shape("=Sheet1!B3+Sheet1!B5")
    shape_c = fingerprint_formula_shape("=Sheet1!C3+Sheet1!C5")
    assert shape_b.skeleton == shape_c.skeleton
    rebound = specialize_formula_shape(shape_b.skeleton, shape_c.params)
    assert rebound == parse("=Sheet1!C3+Sheet1!C5")


def test_specialize_rejects_arity_mismatch() -> None:
    shape = fingerprint_formula_shape("=Sheet1!A1+Sheet1!B1")
    with pytest.raises(ValueError, match="arity|param|hole"):
        specialize_formula_shape(shape.skeleton, (CellRefNode("Sheet1!A1"),))


def test_specialize_rejects_kind_mismatch() -> None:
    shape = fingerprint_formula_shape("=Sheet1!A1+Sheet1!B1")
    bad_params = (CellRefNode("Sheet1!A1"), RangeNode("Sheet1!A1", "Sheet1!A2"))
    with pytest.raises(ValueError, match="kind"):
        specialize_formula_shape(shape.skeleton, bad_params)


def test_mean_instances_per_shape_ignores_unparseable_count() -> None:
    """`formula_nodes` is successes-only; do not subtract `unparseable` again."""
    summary = FormulaShapeSummary(
        formula_nodes=3,
        distinct_normalized_formulas=3,
        distinct_shapes=2,
        unparseable=1,
        shape_counts=(("B(+,$CELL,$CELL)", 2), ("$CELL", 1)),
    )
    assert summary.mean_instances_per_shape == 1.5


def test_summarize_normalized_formulas_counts_unparseable_separately() -> None:
    summary, parseable = summarize_normalized_formulas(
        [
            "=Sheet1!A1+Sheet1!B1",
            "=Sheet1!A2+Sheet1!B2",
            "=1+",  # parse error
            "=$CELL",  # bare ref — not sheet-qualified
        ]
    )
    assert summary.unparseable == 2
    assert summary.formula_nodes == 2
    assert summary.distinct_shapes == 1
    assert summary.mean_instances_per_shape == 2.0
    assert parseable == ["=Sheet1!A1+Sheet1!B1", "=Sheet1!A2+Sheet1!B2"]


@pytest.mark.skipif(not _FFV2.is_file(), reason="ffv2.xlsx fixture missing")
def test_summarize_ffv2_collapses_autofill_column_to_one_shape() -> None:
    graph = create_dependency_graph(_FFV2, ["Sheet1!B18:Q18"], load_values=False)
    summary = summarize_formula_shapes(graph)
    assert summary.formula_nodes == 16
    assert summary.distinct_normalized_formulas == 16
    assert summary.distinct_shapes == 1
    assert summary.shapes_per_formula_string < 0.1


@pytest.mark.skipif(not _TACO.is_file(), reason="taco_patterns.xlsx fixture missing")
def test_summarize_taco_patterns_collapses_strings_to_few_shapes() -> None:
    graph = create_dependency_graph(_TACO, list(_TACO_TARGETS), load_values=False)
    summary = summarize_formula_shapes(graph)
    assert summary.formula_nodes == 24
    assert summary.distinct_normalized_formulas == 24
    assert summary.distinct_shapes == 4
    assert summary.mean_instances_per_shape == 6.0
