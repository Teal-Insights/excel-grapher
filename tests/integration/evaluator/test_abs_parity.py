"""ABS: evaluator and generated export runtime agree on synthetic graphs (integration).

Live Excel parity for ``ABS`` lives in ``test_abs_excel_parity.py`` (slow, run-if-available).
"""

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
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


def test_abs_parity_literal_nested_and_cell_ref() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, -4),
        _make_node("S!B1", "=ABS(-3)", None),
        _make_node("S!B2", "=ABS(S!A1)", None),
        _make_node("S!B3", "=SUM(ABS(S!A1),10)", None),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2", "S!B3"])
    assert result.generated_results["S!B1"] == 3.0
    assert result.generated_results["S!B2"] == 4.0
    assert result.generated_results["S!B3"] == 14.0
    assert "xl_abs" in result.generated_code


def test_abs_parity_sigma_band_pattern() -> None:
    """Mirror ``normdist_sigma_band`` IF/ABS nesting from the sandbox workbook."""
    graph = _make_graph(
        _make_node("S!J1", None, 0.5),
        _make_node(
            "S!M1",
            '=IF(ABS(S!J1)<=1,"Within 1σ",IF(ABS(S!J1)<=2,"Within 2σ","Outlier >2σ"))',
            None,
        ),
        _make_node("S!J2", None, 2.5),
        _make_node(
            "S!M2",
            '=IF(ABS(S!J2)<=1,"Within 1σ",IF(ABS(S!J2)<=2,"Within 2σ","Outlier >2σ"))',
            None,
        ),
    )

    result = assert_codegen_matches_evaluator(graph, ["S!M1", "S!M2"])
    assert result.generated_results["S!M1"] == "Within 1σ"
    assert result.generated_results["S!M2"] == "Outlier >2σ"
