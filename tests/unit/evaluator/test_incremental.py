"""Tests for incremental computation and cache invalidation."""


from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.core.address_keys import parse_address


def _make_node(address: str, formula: str | None, value: object) -> Node:
    """Helper to create a Node from a sheet-qualified address."""
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
    """Helper to create a DependencyGraph from nodes."""
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


# --- durable leaf update via graph.set_node_value ---


def test_graph_set_node_value_updates_node() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.set_node_value("S!A1", 20)
    node = graph.get_node("S!A1")
    assert node is not None
    assert node.value == 20


def test_reevaluation_after_graph_set_node_value_uses_new_value() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)
        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


def test_graph_set_node_value_dependents_re_evaluate() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!C1", "=S!B1+1", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!C1", "S!B1")

    with FormulaEvaluator(graph) as ev:
        first = ev.evaluate(["S!C1"])
        assert first["S!C1"] == 21.0

        graph.set_node_value("S!A1", 20)
        second = ev.evaluate(["S!C1"])
        assert second["S!C1"] == 41.0


def test_graph_set_node_value_does_not_affect_unrelated_cells() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(graph) as ev:
        first = ev.evaluate(["S!B1", "S!B2"])
        assert first == {"S!B1": 20.0, "S!B2": 15.0}

        graph.set_node_value("S!A1", 20)
        second = ev.evaluate(["S!B1", "S!B2"])
        assert second == {"S!B1": 40.0, "S!B2": 15.0}


# --- auto_detect_changes tests ---


def test_auto_detect_changes_detects_mutated_leaf() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=True) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)

        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


def test_auto_detect_changes_false_ignores_durable_leaf_update() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=False) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)

        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 20.0


# --- eager_invalidation tests ---


def test_eager_invalidation_checks_all_leaves_upfront() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=True) as ev:
        ev.evaluate(["S!B1", "S!B2"])

        graph.set_node_value("S!A1", 1)
        graph.set_node_value("S!A2", 2)

        ev.evaluate(["S!B1"])
        # Eager mode checks all leaves up-front, so A2's change is observed
        # even though we only evaluate B1.
        assert ev._cache.get("S!A2") in (None, 2)  # noqa: SLF001


def test_lazy_invalidation_only_checks_visited_leaves() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=False) as ev:
        ev.evaluate(["S!B1", "S!B2"])
        assert ev._cache["S!B2"] == 15.0  # noqa: SLF001

        graph.set_node_value("S!A2", 100)

        ev.evaluate(["S!B1"])
        # B2's cached value should still be stale in lazy mode
        assert ev._cache.get("S!B2") == 15.0  # noqa: SLF001


def test_lazy_invalidation_detects_changes_in_evaluation_path() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=False) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)

        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0
