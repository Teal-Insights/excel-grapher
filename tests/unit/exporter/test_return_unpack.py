"""Tests for optional return-line unpacking in codegen."""

from __future__ import annotations

import ast
from typing import cast

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.return_unpack import unpack_return_expression


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


class TestUnpackReturnExpression:
    def test_no_nested_calls_unchanged(self) -> None:
        expr = "xl_cell(ctx, 'Sheet1!A1')"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == []
        assert return_expr == expr
        assert counter == 0

    def test_single_nested_call(self) -> None:
        expr = "xl_add(xl_eval(ctx, 'Sheet1!C5', cell_sheet1_c5), 1.0)"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == ["_t1 = xl_eval(ctx, 'Sheet1!C5', cell_sheet1_c5)"]
        assert return_expr == "xl_add(_t1, 1.0)"
        assert counter == 1

    def test_multiple_args_in_call_order(self) -> None:
        expr = "xl_sum(xl_cell(ctx, 'Sheet1!A1'), xl_cell(ctx, 'Sheet1!A2'))"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == [
            "_t1 = xl_cell(ctx, 'Sheet1!A1')",
            "_t2 = xl_cell(ctx, 'Sheet1!A2')",
        ]
        assert return_expr == "xl_sum(_t1, _t2)"
        assert counter == 2

    def test_deeply_nested_calls(self) -> None:
        expr = "xl_foo(xl_bar(xl_baz()), xl_qux())"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == [
            "_t1 = xl_baz()",
            "_t2 = xl_bar(_t1)",
            "_t3 = xl_qux()",
        ]
        assert return_expr == "xl_foo(_t2, _t3)"
        assert counter == 3

    def test_respects_existing_temp_counter(self) -> None:
        expr = "xl_add(xl_cell(ctx, 'Sheet1!A1'), 1.0)"
        statements, return_expr, counter = unpack_return_expression(expr, 1)
        assert statements == ["_t2 = xl_cell(ctx, 'Sheet1!A1')"]
        assert return_expr == "xl_add(_t2, 1.0)"
        assert counter == 2

    def test_skips_calls_inside_lambda(self) -> None:
        expr = "xl_iferror(lambda: (xl_foo()), lambda: (xl_bar()))"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == []
        assert "lambda:" in return_expr
        assert "xl_foo()" in return_expr
        assert "xl_bar()" in return_expr
        assert counter == 0

    def test_skips_calls_inside_ifexp_branches(self) -> None:
        expr = "((xl_true()) if cond else (xl_false()))"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == []
        assert "xl_true()" in return_expr
        assert "xl_false()" in return_expr
        assert "if cond else" in return_expr
        assert counter == 0

    def test_hoists_calls_in_ifexp_test_only(self) -> None:
        expr = "((xl_a()) if xl_bool(xl_cell(ctx, 'Sheet1!A1')) else (xl_b()))"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == ["_t1 = xl_cell(ctx, 'Sheet1!A1')"]
        assert "xl_a()" in return_expr
        assert "xl_b()" in return_expr
        assert "xl_bool(_t1)" in return_expr
        assert counter == 1

    def test_skips_calls_inside_boolop(self) -> None:
        expr = "xl_guard(xl_a() or xl_b())"
        statements, return_expr, counter = unpack_return_expression(expr, 0)
        assert statements == []
        assert return_expr == expr
        assert counter == 0


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
        assert "return xl_number(_t1) + xl_number(1.0)" in code
        compile(code, "<string>", "exec")

    def test_cycle_cell_unpacks_xl_eval(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!B5", "=Sheet1!C5+1", None),
            _make_node("Sheet1!C5", "=Sheet1!B5+1", None),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code_b5 = gen._emit_cell("Sheet1!B5")
        assert "_t1 = xl_eval(ctx, 'Sheet1!C5', cell_sheet1_c5)" in code_b5
        assert "return xl_number(_t1) + xl_number(1.0)" in code_b5

    def test_if_branches_remain_lazy(self) -> None:
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!B1", "=IF(Sheet1!A1, Sheet1!C1, Sheet1!D1)", None),
            _make_node("Sheet1!C1", None, 10.0),
            _make_node("Sheet1!D1", None, 20.0),
        )
        gen = CodeGenerator(graph, unpack_return=True)
        code = gen._emit_cell("Sheet1!B1")
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
        compute_all = cast(object, ns["compute_all"])
        assert callable(compute_all)
        generated_results = compute_all()
        assert generated_results == evaluator_results
