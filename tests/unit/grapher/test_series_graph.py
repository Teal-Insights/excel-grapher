"""Series-level quotient: one node per bound series, cross-series edges merged."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef as GuardCellRef
from excel_grapher.grapher.guard import Compare, Literal
from excel_grapher.grapher.node import Node
from excel_grapher.grapher.series_graph import SeriesGraph, to_series_graph


@dataclass(frozen=True, slots=True)
class _FakeSeries:
    series_id: str
    direction: str = "internal"
    layout: str = "series"
    cells: tuple[str, ...] = ()


@dataclass
class _FakeOccupancy:
    order: tuple[str, ...]
    address_to_id: dict[str, str]
    series: dict[str, _FakeSeries]

    def series_id_for(self, address: str) -> str | None:
        return self.address_to_id.get(address)

    def get(self, series_id: str) -> _FakeSeries:
        return self.series[series_id]


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


def _two_series_model() -> tuple[DependencyGraph, _FakeOccupancy]:
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
    occupancy = _FakeOccupancy(
        order=("inputs", "calc", "outputs"),
        address_to_id={
            "Inputs!A1": "inputs",
            "Inputs!B1": "inputs",
            "Calc!A1": "calc",
            "Calc!B1": "calc",
            "Outputs!A1": "outputs",
        },
        series={
            "inputs": _FakeSeries("inputs", direction="input", cells=("Inputs!A1", "Inputs!B1")),
            "calc": _FakeSeries("calc", cells=("Calc!A1", "Calc!B1")),
            "outputs": _FakeSeries("outputs", direction="output", cells=("Outputs!A1",)),
        },
    )
    return graph, occupancy


def test_to_series_graph_emits_one_node_per_series_in_bindings_order() -> None:
    graph, occupancy = _two_series_model()
    series_graph = to_series_graph(graph, occupancy)

    assert [node.series_id for node in series_graph.nodes] == ["inputs", "calc", "outputs"]
    by_id = {node.series_id: node for node in series_graph.nodes}
    assert by_id["inputs"].cell_count == 2
    assert by_id["inputs"].formula_count == 0
    assert by_id["inputs"].leaf_count == 2
    assert by_id["inputs"].direction == "input"
    assert by_id["inputs"].sheet == "Inputs"
    assert by_id["calc"].cell_count == 2
    assert by_id["calc"].formula_count == 2
    assert by_id["calc"].leaf_count == 0
    assert by_id["outputs"].cell_count == 1
    assert by_id["outputs"].direction == "output"


def test_to_series_graph_consolidates_cross_series_edges_and_drops_intra_series() -> None:
    graph, occupancy = _two_series_model()
    series_graph = to_series_graph(graph, occupancy)

    edges = {(edge.source, edge.target): edge for edge in series_graph.edges}
    assert set(edges) == {("calc", "inputs"), ("outputs", "calc")}
    assert edges[("calc", "inputs")].edge_count == 2
    assert edges[("calc", "inputs")].guarded_count == 0
    assert edges[("outputs", "calc")].edge_count == 1
    assert ("calc", "calc") not in edges


def test_to_series_graph_can_keep_intra_series_self_loops() -> None:
    graph, occupancy = _two_series_model()
    series_graph = to_series_graph(graph, occupancy, include_self_loops=True)

    edges = {(edge.source, edge.target): edge for edge in series_graph.edges}
    assert ("calc", "calc") in edges
    assert edges[("calc", "calc")].edge_count == 1


def test_to_series_graph_keeps_isolated_series() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Notes", "A", 1))
    graph.add_node(_leaf("Data", "A", 1))
    graph.add_node(_formula("Data", "B", 1, "=A1"))
    graph.add_edge("Data!B1", "Data!A1")
    occupancy = _FakeOccupancy(
        order=("notes", "data"),
        address_to_id={"Notes!A1": "notes", "Data!A1": "data", "Data!B1": "data"},
        series={
            "notes": _FakeSeries("notes", direction="constant", cells=("Notes!A1",)),
            "data": _FakeSeries("data", cells=("Data!A1", "Data!B1")),
        },
    )

    series_graph = to_series_graph(graph, occupancy)
    assert [node.series_id for node in series_graph.nodes] == ["notes", "data"]
    assert series_graph.edges == ()


def test_to_series_graph_counts_guarded_cell_edges() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Inputs", "A", 1))
    graph.add_node(_leaf("Inputs", "C", 1))
    graph.add_node(_formula("Calc", "A", 1, "=IF(Inputs!C1,Inputs!A1,0)"))
    graph.add_node(_formula("Calc", "B", 1, "=Inputs!A1"))
    guard = Compare(left=GuardCellRef(key="Inputs!C1"), op="=", right=Literal(True))
    graph.add_edge("Calc!A1", "Inputs!A1", guard=guard)
    graph.add_edge("Calc!B1", "Inputs!A1")
    occupancy = _FakeOccupancy(
        order=("inputs", "calc"),
        address_to_id={
            "Inputs!A1": "inputs",
            "Inputs!C1": "inputs",
            "Calc!A1": "calc",
            "Calc!B1": "calc",
        },
        series={
            "inputs": _FakeSeries("inputs", direction="input", cells=("Inputs!A1", "Inputs!C1")),
            "calc": _FakeSeries("calc", cells=("Calc!A1", "Calc!B1")),
        },
    )

    series_graph = to_series_graph(graph, occupancy)
    assert len(series_graph.edges) == 1
    edge = series_graph.edges[0]
    assert (edge.source, edge.target) == ("calc", "inputs")
    assert edge.edge_count == 2
    assert edge.guarded_count == 1


def test_to_series_graph_skips_unbound_endpoints() -> None:
    graph = DependencyGraph()
    graph.add_node(_formula("Calc", "A", 1, "=Missing!A1"))
    graph.add_node(_leaf("Missing", "A", 1))
    graph.add_edge("Calc!A1", "Missing!A1")
    occupancy = _FakeOccupancy(
        order=("calc",),
        address_to_id={"Calc!A1": "calc"},
        series={"calc": _FakeSeries("calc", cells=("Calc!A1",))},
    )

    series_graph = to_series_graph(graph, occupancy)
    assert [node.series_id for node in series_graph.nodes] == ["calc"]
    assert series_graph.edges == ()


def test_to_series_graph_keeps_both_directions_of_a_series_cycle() -> None:
    """Series quotients invent cycles; this visualizer does not rank them away."""
    graph = DependencyGraph()
    graph.add_node(_formula("S", "A", 1, "=100"))
    graph.add_node(_formula("S", "A", 2, "=A1+B2"))
    graph.add_node(_formula("S", "B", 2, "=A1*0.02"))
    graph.add_edge("S!A2", "S!A1")
    graph.add_edge("S!A2", "S!B2")
    graph.add_edge("S!B2", "S!A1")
    occupancy = _FakeOccupancy(
        order=("debt", "adjustment"),
        address_to_id={"S!A1": "debt", "S!A2": "debt", "S!B2": "adjustment"},
        series={
            "debt": _FakeSeries("debt", direction="output", cells=("S!A1", "S!A2")),
            "adjustment": _FakeSeries("adjustment", cells=("S!B2",)),
        },
    )

    series_graph = to_series_graph(graph, occupancy)
    edges = {(edge.source, edge.target): edge for edge in series_graph.edges}
    assert set(edges) == {("debt", "adjustment"), ("adjustment", "debt")}
    assert edges[("debt", "adjustment")].edge_count == 1
    assert edges[("adjustment", "debt")].edge_count == 1


def test_to_series_graph_marks_mixed_sheet_series() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("A", "A", 1))
    graph.add_node(_leaf("B", "A", 1))
    occupancy = _FakeOccupancy(
        order=("split",),
        address_to_id={"A!A1": "split", "B!A1": "split"},
        series={"split": _FakeSeries("split", cells=("A!A1", "B!A1"))},
    )

    series_graph = to_series_graph(graph, occupancy)
    assert series_graph.nodes[0].sheet == "mixed"


def test_to_graphviz_renders_consolidated_series_edges() -> None:
    from excel_grapher.grapher import to_graphviz

    graph, occupancy = _two_series_model()
    dot = to_graphviz(to_series_graph(graph, occupancy), rankdir="LR")
    assert "digraph series_dependencies" in dot
    assert "rankdir=LR" in dot
    assert '"inputs"' in dot
    assert '"calc"' in dot
    assert '"outputs"' in dot
    assert '"calc" -> "inputs"' in dot
    assert "label=" in dot
    assert '"calc" -> "calc"' not in dot
    assert "Calc!A1" not in dot


def test_to_graphviz_dashes_fully_guarded_series_edges() -> None:
    from excel_grapher.grapher import to_graphviz

    graph = DependencyGraph()
    graph.add_node(_leaf("Inputs", "A", 1))
    graph.add_node(_formula("Calc", "A", 1, "=IF(1,Inputs!A1,0)"))
    guard = Compare(left=Literal(1), op="=", right=Literal(1))
    graph.add_edge("Calc!A1", "Inputs!A1", guard=guard)
    occupancy = _FakeOccupancy(
        order=("inputs", "calc"),
        address_to_id={"Inputs!A1": "inputs", "Calc!A1": "calc"},
        series={
            "inputs": _FakeSeries("inputs", direction="input", cells=("Inputs!A1",)),
            "calc": _FakeSeries("calc", cells=("Calc!A1",)),
        },
    )

    dot = to_graphviz(to_series_graph(graph, occupancy))
    assert '"calc" -> "inputs"' in dot
    assert "style=dashed" in dot


def test_to_mermaid_renders_consolidated_series_edges() -> None:
    from excel_grapher.grapher import to_mermaid

    graph, occupancy = _two_series_model()
    mm = to_mermaid(to_series_graph(graph, occupancy))
    assert mm.startswith("flowchart TD")
    assert "inputs" in mm
    assert "calc" in mm
    assert "outputs" in mm
    assert "calc --> inputs" in mm or "calc -->|" in mm
    assert "Calc!A1" not in mm


def test_to_networkx_uses_series_ids_and_edge_weights() -> None:
    from excel_grapher.grapher import to_networkx

    graph, occupancy = _two_series_model()
    nx_graph = to_networkx(to_series_graph(graph, occupancy))
    assert set(nx_graph.nodes) == {"inputs", "calc", "outputs"}
    assert nx_graph.nodes["inputs"]["cell_count"] == 2
    assert nx_graph.nodes["inputs"]["direction"] == "input"
    assert nx_graph.edges[("calc", "inputs")]["weight"] == 2
    assert nx_graph.edges[("calc", "inputs")]["edge_count"] == 2
    assert ("calc", "calc") not in nx_graph.edges


def test_series_graph_is_public_type() -> None:
    assert SeriesGraph.__name__ == "SeriesGraph"
