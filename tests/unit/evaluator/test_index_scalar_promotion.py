"""INDEX single-cell results must scalar-promote in value contexts (issue #264)."""

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


def test_text_index_match_promotes_single_cell_index_to_scalar() -> None:
    """``TEXT(INDEX(..., MATCH(...)), ...)`` formats a scalar, not a stringified array."""
    graph = _make_graph(
        _make_node("PL!K5", None, "PRD-001"),
        _make_node("PL!A5", None, "PRD-001"),
        _make_node("PL!A6", None, "PRD-002"),
        _make_node("PL!E5", None, 1499.0),
        _make_node("PL!E6", None, 999.0),
        _make_node(
            "PL!K16",
            '=TEXT(INDEX(PL!E5:PL!E6,MATCH(PL!K5,PL!A5:PL!A6,0)),"0.00")',
            None,
        ),
    )
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("PL!K16") == "1499.00"


def test_numbervalue_text_index_match_returns_numeric_price(tmp_path) -> None:
    """``NUMBERVALUE(TEXT(INDEX(...)))`` returns the looked-up price (K16 shape)."""
    workbook = tmp_path / "numbervalue_lookup.xlsx"
    writer = xlsxwriter.Workbook(workbook)
    worksheet = writer.add_worksheet("Product Lookup")
    worksheet.write_string(4, 10, "PRD-001")
    for row, sku in enumerate(["PRD-001", "PRD-002"], start=5):
        worksheet.write_string(row - 1, 0, sku)
        worksheet.write_number(row - 1, 4, 1499.0)
    worksheet.write_formula(
        15,
        10,
        '=IFERROR(NUMBERVALUE(TEXT(INDEX($E$5:$E$19,MATCH($K$5,$A$5:$A$19,0)),"0.00"),".",","),"N/A")',
        None,
        1499,
    )
    writer.close()

    graph = create_dependency_graph(
        workbook,
        ["Product Lookup!K16"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as evaluator:
        assert evaluator.evaluate("Product Lookup!K16") == 1499.0


def test_numbervalue_text_index_match_eval_codegen_parity(tmp_path) -> None:
    """Evaluator and export agree on the K16 ``NUMBERVALUE(TEXT(INDEX(...)))`` chain."""
    workbook = tmp_path / "numbervalue_lookup_parity.xlsx"
    writer = xlsxwriter.Workbook(workbook)
    worksheet = writer.add_worksheet("Product Lookup")
    worksheet.write_string(4, 10, "PRD-001")
    for row, sku in enumerate(["PRD-001", "PRD-002"], start=5):
        worksheet.write_string(row - 1, 0, sku)
        worksheet.write_number(row - 1, 4, 1499.0)
    worksheet.write_formula(
        15,
        10,
        '=IFERROR(NUMBERVALUE(TEXT(INDEX($E$5:$E$19,MATCH($K$5,$A$5:$A$19,0)),"0.00"),".",","),"N/A")',
        None,
        1499,
    )
    writer.close()

    graph = create_dependency_graph(
        workbook,
        ["Product Lookup!K16"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    result = assert_codegen_matches_evaluator(graph, ["Product Lookup!K16"])
    assert result.evaluator_results["Product Lookup!K16"] == 1499.0
    assert result.generated_results["Product Lookup!K16"] == 1499.0
