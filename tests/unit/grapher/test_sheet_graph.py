"""Sheet-level quotient: one node per sheet, cross-sheet edges merged."""

from __future__ import annotations

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef as GuardCellRef
from excel_grapher.grapher.guard import Compare, Literal
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.sheet_graph import SheetGraph, to_sheet_graph


def _leaf(sheet: str, col: str, row: int, value: object = 0) -> Node:
    return Node(sheet=sheet, column=col, row=row, value=value, is_leaf=True)


def _formula(sheet: str, col: str, row: int, formula: str) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        is_leaf=False,
    )


def _two_sheet_model() -> DependencyGraph:
    graph = DependencyGraph(sheet_order=["Inputs", "Calc", "Outputs"])
    graph.add_node(_leaf("Inputs", "A", 1, 2))
    graph.add_node(_leaf("Inputs", "B", 1, 3))
    graph.add_node(_formula("Calc", "A", 1, "=Inputs!A1+Inputs!B1"))
    graph.add_node(_formula("Calc", "B", 1, "=A1*2"))
    graph.add_node(_formula("Outputs", "A", 1, "=Calc!A1"))
    graph.add_edge("Calc!A1", "Inputs!A1")
    graph.add_edge("Calc!A1", "Inputs!B1")
    graph.add_edge("Calc!B1", "Calc!A1")
    graph.add_edge("Outputs!A1", "Calc!A1")
    return graph


def test_to_sheet_graph_emits_one_node_per_sheet_in_workbook_order() -> None:
    sheet_graph = to_sheet_graph(_two_sheet_model())

    assert [node.name for node in sheet_graph.nodes] == ["Inputs", "Calc", "Outputs"]
    by_name = {node.name: node for node in sheet_graph.nodes}
    assert by_name["Inputs"].cell_count == 2
    assert by_name["Inputs"].formula_count == 0
    assert by_name["Inputs"].leaf_count == 2
    assert by_name["Calc"].cell_count == 2
    assert by_name["Calc"].formula_count == 2
    assert by_name["Calc"].leaf_count == 0
    assert by_name["Outputs"].cell_count == 1


def test_to_sheet_graph_consolidates_cross_sheet_edges_and_drops_intra_sheet() -> None:
    sheet_graph = to_sheet_graph(_two_sheet_model())

    edges = {(edge.source, edge.target): edge for edge in sheet_graph.edges}
    assert set(edges) == {("Calc", "Inputs"), ("Outputs", "Calc")}
    assert edges[("Calc", "Inputs")].edge_count == 2
    assert edges[("Calc", "Inputs")].guarded_count == 0
    assert edges[("Outputs", "Calc")].edge_count == 1
    assert ("Calc", "Calc") not in edges


def test_to_sheet_graph_can_keep_intra_sheet_self_loops() -> None:
    sheet_graph = to_sheet_graph(_two_sheet_model(), include_self_loops=True)

    edges = {(edge.source, edge.target): edge for edge in sheet_graph.edges}
    assert ("Calc", "Calc") in edges
    assert edges[("Calc", "Calc")].edge_count == 1


def test_to_sheet_graph_keeps_isolated_sheets() -> None:
    graph = DependencyGraph(sheet_order=["Notes", "Data"])
    graph.add_node(_leaf("Notes", "A", 1))
    graph.add_node(_leaf("Data", "A", 1))
    graph.add_node(_formula("Data", "B", 1, "=A1"))
    graph.add_edge("Data!B1", "Data!A1")

    sheet_graph = to_sheet_graph(graph)
    assert [node.name for node in sheet_graph.nodes] == ["Notes", "Data"]
    assert sheet_graph.edges == ()


def test_to_sheet_graph_counts_guarded_cell_edges() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Inputs", "A", 1))
    graph.add_node(_leaf("Inputs", "C", 1))
    graph.add_node(_formula("Calc", "A", 1, "=IF(Inputs!C1,Inputs!A1,0)"))
    graph.add_node(_formula("Calc", "B", 1, "=Inputs!A1"))
    guard = Compare(left=GuardCellRef(key="Inputs!C1"), op="=", right=Literal(True))
    graph.add_edge("Calc!A1", "Inputs!A1", guard=guard)
    graph.add_edge("Calc!B1", "Inputs!A1")

    sheet_graph = to_sheet_graph(graph)
    assert len(sheet_graph.edges) == 1
    edge = sheet_graph.edges[0]
    assert (edge.source, edge.target) == ("Calc", "Inputs")
    assert edge.edge_count == 2
    assert edge.guarded_count == 1


def test_to_sheet_graph_skips_dangling_endpoints() -> None:
    graph = DependencyGraph()
    graph.add_node(_formula("Calc", "A", 1, "=Missing!A1"))
    graph.add_edge("Calc!A1", "Missing!A1")

    sheet_graph = to_sheet_graph(graph)
    assert [node.name for node in sheet_graph.nodes] == ["Calc"]
    assert sheet_graph.edges == ()


def test_to_graphviz_renders_consolidated_sheet_edges() -> None:
    from excel_grapher.grapher import to_graphviz

    dot = to_graphviz(to_sheet_graph(_two_sheet_model()), rankdir="LR")
    assert "digraph sheet_dependencies" in dot
    assert "rankdir=LR" in dot
    assert '"Inputs"' in dot
    assert '"Calc"' in dot
    assert '"Outputs"' in dot
    assert '"Calc" -> "Inputs"' in dot
    assert "label=" in dot
    assert '"Calc" -> "Calc"' not in dot
    assert "Sheet1!" not in dot
    assert "Calc!A1" not in dot


def test_to_graphviz_dashes_fully_guarded_sheet_edges() -> None:
    from excel_grapher.grapher import to_graphviz

    graph = DependencyGraph()
    graph.add_node(_leaf("Inputs", "A", 1))
    graph.add_node(_formula("Calc", "A", 1, "=IF(1,Inputs!A1,0)"))
    guard = Compare(left=Literal(1), op="=", right=Literal(1))
    graph.add_edge("Calc!A1", "Inputs!A1", guard=guard)

    dot = to_graphviz(to_sheet_graph(graph))
    assert '"Calc" -> "Inputs"' in dot
    assert "style=dashed" in dot


def test_to_mermaid_renders_consolidated_sheet_edges() -> None:
    from excel_grapher.grapher import to_mermaid

    mm = to_mermaid(to_sheet_graph(_two_sheet_model()))
    assert mm.startswith("flowchart TD")
    assert "Inputs" in mm
    assert "Calc" in mm
    assert "Outputs" in mm
    assert "Calc --> Inputs" in mm or "Calc -->|" in mm
    assert "Calc!A1" not in mm


def test_to_networkx_uses_sheet_names_and_edge_weights() -> None:
    from excel_grapher.grapher import to_networkx

    nx_graph = to_networkx(to_sheet_graph(_two_sheet_model()))
    assert set(nx_graph.nodes) == {"Inputs", "Calc", "Outputs"}
    assert nx_graph.nodes["Inputs"]["cell_count"] == 2
    assert nx_graph.edges[("Calc", "Inputs")]["weight"] == 2
    assert nx_graph.edges[("Calc", "Inputs")]["edge_count"] == 2
    assert ("Calc", "Calc") not in nx_graph.edges


def test_sheet_graph_is_public_type() -> None:
    assert SheetGraph.__name__ == "SheetGraph"
