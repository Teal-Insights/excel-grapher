"""Shared evaluator ↔ export parity scenarios for reuse across harness tests."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address


@dataclass(frozen=True, slots=True)
class ParityScenario:
    name: str
    graph: DependencyGraph
    targets: list[str]
    rtol: float = 0.0
    atol: float = 0.0
    blank_ranges: tuple[str, ...] | None = None


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


def _array_graph(formula: str) -> DependencyGraph:
    nodes = [_make_node("S!Z1", formula, None)]
    for col in "ABCD":
        for row in (1, 2, 3):
            nodes.append(_make_node(f"S!{col}{row}", None, float(row)))
    return _make_graph(*nodes)


def parity_scenarios() -> list[ParityScenario]:
    golden = _make_graph(
        _make_node("S!A1", None, "0"),
        _make_node("S!A2", None, 0),
        _make_node("S!A3", None, "FALSE"),
        _make_node("S!B1", "=S!A1=S!A2", None),
        _make_node("S!B2", "=IF(S!A3,1,2)", None),
        _make_node("S!B3", "=IFERROR(1/0,99)", None),
        _make_node("S!B4", "=SUM(S!B2,S!B3)", None),
        _make_node("S!C1", "=SUM(S!B1:S!B4)", None),
    )

    return [
        ParityScenario(
            name="golden_mixed",
            graph=golden,
            targets=["S!B1", "S!B2", "S!B3", "S!B4", "S!C1"],
        ),
        ParityScenario(
            name="simple_arithmetic",
            graph=_make_graph(
                _make_node("S!A1", None, 10),
                _make_node("S!A2", None, 5),
                _make_node("S!B1", "=S!A1+S!A2", None),
            ),
            targets=["S!B1"],
        ),
        ParityScenario(
            name="xlookup",
            graph=_make_graph(
                _make_node("S!A1", None, 1),
                _make_node("S!A2", None, 2),
                _make_node("S!A3", None, 3),
                _make_node("S!B1", None, "a"),
                _make_node("S!B2", None, "b"),
                _make_node("S!B3", None, "c"),
                _make_node("S!C1", "=_xlfn.XLOOKUP(2,S!A1:S!A3,S!B1:S!B3)", None),
            ),
            targets=["S!C1"],
        ),
        ParityScenario(
            name="array_range_add",
            graph=_array_graph("=S!A1:A3+S!B1:B3"),
            targets=["S!Z1"],
        ),
        ParityScenario(
            name="nested_array_sumproduct",
            graph=_array_graph("=SUM((S!A1:A3+S!B1:B3)*(S!C1:C3-S!D1:D3))"),
            targets=["S!Z1"],
        ),
        ParityScenario(
            name="row_with_offset",
            graph=_make_graph(
                _make_node("S!B3", None, 10),
                _make_node("S!C1", None, 2),
                _make_node("S!C2", None, 0),
                _make_node("S!A1", "=ROW(S!B3)", None),
                _make_node("S!A2", "=ROW(OFFSET(S!B3, S!C1, S!C2))", None),
            ),
            targets=["S!A1", "S!A2"],
        ),
        ParityScenario(
            name="if_cycle",
            graph=_make_graph(
                _make_node("S!A1", "=IF(S!B1, 0, 1)", None),
                _make_node("S!B1", "=S!A1", None),
            ),
            targets=["S!A1"],
        ),
    ]
