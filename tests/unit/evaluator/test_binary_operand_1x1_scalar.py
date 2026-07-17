"""1x1 range operands of binary ops must resolve to scalars (issue #421).

Excel collapses a single-cell reference (including INDEX returning one cell)
to its value in scalar context. Materializing 1x1 ranges as arrays makes
``IF(INDEX(...)=\"Yes\", ...)`` return ``#VALUE!``, and ``IFERROR`` then
silently substitutes the fallback.
"""

from __future__ import annotations

import xlsxwriter

from excel_grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def test_if_index_equality_uses_scalar_condition() -> None:
    """``IF(INDEX(...)=\"Yes\", 1, 2)`` returns 1, not ``#VALUE!``."""
    graph = _make_graph(
        _make_node("S!A1", None, "Yes"),
        _make_node("S!A2", None, "No"),
        _make_node("S!B1", '=IF(INDEX(S!A1:S!A2,1,1)="Yes",1,2)', None),
    )
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("S!B1") == 1


def test_iferror_index_flag_does_not_take_fallback() -> None:
    """``IFERROR(IF(INDEX(...)=\"Yes\", ...), fallback)`` keeps the true branch."""
    graph = _make_graph(
        _make_node("S!A1", None, "Yes"),
        _make_node("S!A2", None, "No"),
        _make_node(
            "S!B2",
            '=IFERROR(IF(INDEX(S!A1:S!A2,1,1)="Yes","Yes",""),"")',
            None,
        ),
    )
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("S!B2") == "Yes"


def test_index_concat_returns_scalar_string() -> None:
    """``INDEX(...)&INDEX(...)`` concatenates cell values into one string."""
    graph = _make_graph(
        _make_node("S!A1", None, "Yes"),
        _make_node("S!A2", None, "No"),
        _make_node("S!B3", "=INDEX(S!A1:S!A2,2,1)&INDEX(S!A1:S!A2,1,1)", None),
    )
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("S!B3") == "NoYes"


def test_index_binary_ops_eval_codegen_parity(tmp_path) -> None:
    """Evaluator and export agree on INDEX 1x1 binary-op formulas from issue #421."""
    workbook = tmp_path / "mcve_1x1_ranges.xlsx"
    writer = xlsxwriter.Workbook(workbook)
    worksheet = writer.add_worksheet("S")
    worksheet.write_string(0, 0, "Yes")
    worksheet.write_string(1, 0, "No")
    worksheet.write_formula(0, 1, '=IF(INDEX(A1:A2,1,1)="Yes",1,2)', None, 1)
    worksheet.write_formula(
        1,
        1,
        '=IFERROR(IF(INDEX(A1:A2,1,1)="Yes","Yes",""),"")',
        None,
        "Yes",
    )
    worksheet.write_formula(2, 1, "=INDEX(A1:A2,2,1)&INDEX(A1:A2,1,1)", None, "NoYes")
    writer.close()

    cells = ["S!B1", "S!B2", "S!B3"]
    graph = create_dependency_graph(workbook, cells, load_values=True)
    result = assert_codegen_matches_evaluator(graph, cells)
    assert result.evaluator_results == {
        "S!B1": 1,
        "S!B2": "Yes",
        "S!B3": "NoYes",
    }
    assert result.generated_results == result.evaluator_results
