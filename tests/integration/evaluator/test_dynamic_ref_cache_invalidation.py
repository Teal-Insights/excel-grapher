"""Dynamic-ref cache invalidation: evaluator must follow argument-driven resolution shifts.

Reproduces issue #300. With `use_cached_dynamic_refs=True` the resolved target of an
`OFFSET`/`INDIRECT` is frozen as a static dependency edge at graph-build time. When a
dynamic-ref *argument* changes the range the function resolves to at runtime, an
invalidation strategy keyed on those static edges misses changes propagating through the
newly-resolved chain, producing a silently stale result.

The exported runtime (`EvalContext`) records dependencies dynamically at eval time and
self-corrects, so this is also an evaluator <-> export parity check.
"""

from __future__ import annotations

from collections.abc import Callable

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core import CellValue
from excel_grapher.core.address_keys import parse_address
from excel_grapher.runtime.cache import xl_cell
from excel_grapher.runtime.cache_context import EvalContext
from excel_grapher.runtime.offset_runtime import xl_offset


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


def _frozen_cached_offset_graph() -> DependencyGraph:
    """Graph mirroring a `use_cached_dynamic_refs=True` build of an argument-driven OFFSET.

    `S!A1 = OFFSET(S!C1, 0, S!B1)` with cached `S!B1 = 0` freezes the resolved-target edge
    to `S!C1`. `S!D1 = S!X1 * 2` is an independent in-graph chain that the OFFSET resolves
    onto once `S!B1` becomes 1 (column C + 1 -> column D).
    """
    graph = DependencyGraph()
    for node in (
        _make_node("S!X1", None, 5),
        _make_node("S!D1", "=S!X1*2", 10),
        _make_node("S!C1", None, 100),
        _make_node("S!B1", None, 0),
        _make_node("S!A1", "=OFFSET(S!C1, 0, S!B1)", 100),
    ):
        graph.add_node(node)
    # Frozen build-time edges at cached B1 = 0 (resolved target = C1) plus the OFFSET arg.
    graph.add_edge("S!A1", "S!C1")
    graph.add_edge("S!A1", "S!B1")
    # Independent chain that the runtime resolution shifts onto when B1 -> 1.
    graph.add_edge("S!D1", "S!X1")
    return graph


def _evaluator_sequence() -> list[object]:
    """Drive the evaluator through the argument-shift -> dependency-change sequence."""
    graph = _frozen_cached_offset_graph()
    values: list[object] = []
    with FormulaEvaluator(graph) as ev:
        values.append(ev.evaluate("S!A1"))  # B1 = 0 -> resolves C1 = 100
        graph.set_node_value("S!B1", 1)
        values.append(ev.evaluate("S!A1"))  # B1 = 1 -> resolves D1 = X1*2 = 10
        graph.set_node_value("S!X1", 50)
        values.append(ev.evaluate("S!A1"))  # D1 now 100, so A1 must recompute to 100
    return values


def _exported_runtime_resolver() -> Callable[[str], Callable[[EvalContext], CellValue] | None]:
    """Resolver mirroring exported code: dynamic OFFSET + an independent `D1 = X1*2` chain.

    Uses the same `excel_grapher.runtime` primitives the exporter embeds. `xl_offset`
    fetches its resolved target via `xl_cell`, so the runtime dependency edge follows the
    argument-driven shift, exactly as generated standalone code does.
    """

    def _a1(ctx: EvalContext) -> CellValue:
        # Base S!C1 = (sheet "S", row 1, col 3); column offset read from S!B1.
        return xl_offset(ctx, ("S", 1, 3), 0, xl_cell(ctx, "S!B1"))

    def _d1(ctx: EvalContext) -> CellValue:
        factor = xl_cell(ctx, "S!X1")
        assert isinstance(factor, (int, float))
        return factor * 2

    impls: dict[str, Callable[[EvalContext], CellValue]] = {"S!A1": _a1, "S!D1": _d1}
    return impls.get


def _exported_sequence() -> list[object]:
    """Drive the library export runtime through the same input-change sequence."""
    ctx = EvalContext(
        inputs={"S!C1": 100, "S!B1": 0, "S!X1": 5},
        resolver=_exported_runtime_resolver(),
    )
    values: list[object] = []
    values.append(xl_cell(ctx, "S!A1"))  # B1 = 0 -> resolves C1 = 100
    ctx.set_inputs({"S!B1": 1})
    values.append(xl_cell(ctx, "S!A1"))  # B1 = 1 -> resolves D1 = X1*2 = 10
    ctx.set_inputs({"S!X1": 50})
    values.append(xl_cell(ctx, "S!A1"))  # D1 now 100, so A1 must recompute to 100
    return values


def test_evaluator_follows_argument_driven_resolution_shift() -> None:
    """Changing an input through the shifted dependency chain must recompute the OFFSET cell."""
    assert _evaluator_sequence() == [100, 10, 100]


def test_dynamic_ref_invalidation_matches_exported_runtime() -> None:
    """Evaluator and exported runtime agree across the argument-shift sequence (parity)."""
    assert _evaluator_sequence() == _exported_sequence()
