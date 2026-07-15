"""Tests for optional return-line unpacking during formula AST emission."""

from __future__ import annotations

import ast
from collections.abc import Callable
from typing import cast

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator


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
    for col in "AB":
        for row in (1, 2, 3):
            nodes.append(_make_node(f"S!{col}{row}", None, float(row)))
    return _make_graph(*nodes)


class TestEmitCellUnpackReturn:
    def test_disabled_by_default(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!C5", "=Sheet1!B5+1", None),
            _make_node("Sheet1!B5", None, 1.0),
        )
        gen = CodeGenerator(graph)
        code = gen._emit_cell("Sheet1!C5")
        assert "    _t1 = " not in code
        assert "return (xl_number(xl_cell(ctx, 'Sheet1!B5'))" in code

    def test_unpacks_nested_calls_when_enabled(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!C5", "=Sheet1!B5+1", None),
            _make_node("Sheet1!B5", None, 1.0),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code = gen._emit_cell("Sheet1!C5")
        assert "_t1 = xl_cell(ctx, 'Sheet1!B5')" in code
        assert "return (xl_number(_t1) + xl_number(1.0))" in code
        compile(code, "<string>", "exec")

    def test_cycle_cell_unpacks_xl_eval(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!B5", "=Sheet1!C5+1", None),
            _make_node("Sheet1!C5", "=Sheet1!B5+1", None),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code_b5 = gen._emit_cell("Sheet1!B5")
        assert "_t1 = xl_eval(ctx, 'Sheet1!C5', cell_sheet1_c5)" in code_b5
        assert "return (xl_number(_t1) + xl_number(1.0))" in code_b5

    def test_sum_unpacks_both_cell_args(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!A2", None, 2.0),
            _make_node("Sheet1!B1", "=SUM(Sheet1!A1, Sheet1!A2)", None),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code = gen._emit_cell("Sheet1!B1")
        assert "_t1 = xl_cell(ctx, 'Sheet1!A1')" in code
        assert "_t2 = xl_cell(ctx, 'Sheet1!A2')" in code
        assert "return xl_sum(_t1, _t2)" in code

    def test_if_branches_remain_lazy(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!B1", "=IF(Sheet1!A1, Sheet1!C1, Sheet1!D1)", None),
            _make_node("Sheet1!C1", None, 10.0),
            _make_node("Sheet1!D1", None, 20.0),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code = gen._emit_cell("Sheet1!B1")
        assert "    _t2 = xl_cell(ctx, 'Sheet1!C1')" not in code
        assert "    _t2 = xl_cell(ctx, 'Sheet1!D1')" not in code
        tree = ast.parse(code)
        cell_fn = next(node for node in tree.body if isinstance(node, ast.FunctionDef))
        return_node = next(
            stmt for stmt in cell_fn.body if isinstance(stmt, ast.Return) and stmt.value
        )
        assert return_node.value is not None
        return_src = ast.unparse(return_node.value)
        assert "if (" in return_src
        assert "xl_cell(ctx, 'Sheet1!C1')" in return_src
        assert "xl_cell(ctx, 'Sheet1!D1')" in return_src

    def test_iferror_args_remain_inside_lambda(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!B1", "=IFERROR(Sheet1!A1, Sheet1!C1)", None),
            _make_node("Sheet1!C1", None, 99.0),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code = gen._emit_cell("Sheet1!B1")
        assert "    _t2 = xl_cell(ctx, 'Sheet1!C1')" not in code
        assert "lambda: (xl_cell(ctx, 'Sheet1!C1'))" in code

    def test_nested_iferror_inside_sum_hoists_call(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!A1", "=1/0", None),
            _make_node("Sheet1!B1", None, 5.0),
            _make_node("Sheet1!C1", "=SUM(IFERROR(Sheet1!A1, Sheet1!B1))", None),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code = gen._emit_cell("Sheet1!C1")
        assert "_t1 = xl_iferror(" in code
        assert "return xl_sum(_t1)" in code

    def test_array_operator_hoists_range_operands(self) -> None:
        graph = _array_graph("=S!A1:A3+S!B1:B3")
        gen = CodeGenerator(graph, unpack_return=True)
        code = gen._emit_cell("S!Z1")
        assert "_t1 = xl_range(ctx, 'S!A1:A3')" in code
        assert "_t2 = xl_range(ctx, 'S!B1:B3')" in code
        assert "xl_map_arithmetic(" in code
        assert "xl_is_array(" in code
        assert "(_t1, _t2)" in code
        compile(code, "<string>", "exec")


class TestGenerateUnpackReturnParity:
    def test_parity_with_unpack_return_enabled(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!A2", None, 2.0),
            _make_node("Sheet1!B1", "=SUM(Sheet1!A1, Sheet1!A2)", None),
            _make_node("Sheet1!C1", "=Sheet1!B1*2", None),
        )
        targets = ["Sheet1!C1"]
        with FormulaEvaluator(graph) as ev:
            evaluator_results = ev.evaluate(targets)
        code = CodeGenerator(graph, unpack_return=True).generate(targets)
        ns: dict[str, object] = {}
        exec(code, ns)
        compute_all = cast(Callable[[], dict[str, object]], ns["compute_all"])
        generated_results = compute_all()
        assert generated_results == evaluator_results

    def test_if_cycle_parity_with_unpack_return_enabled(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=IF(S!B1, 0, 1)", None),
            _make_node("S!B1", "=S!A1", None),
        )
        targets = ["S!A1"]
        with FormulaEvaluator(graph) as ev:
            evaluator_results = ev.evaluate(targets)
        code = CodeGenerator(graph, unpack_return=True).generate(targets)
        ns = {}
        exec(code, ns)
        compute_all = cast(Callable[[], dict[str, object]], ns["compute_all"])
        generated_results = compute_all()
        assert generated_results == evaluator_results
