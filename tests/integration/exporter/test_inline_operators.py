"""Inlined operators in exported codegen (#316).

Generated scalar formulas use native Python operators with explicit Excel
coercion; array operands still broadcast through compact map helpers.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, cast

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.codegen import CodeGenerator
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


class TestInlineOperatorParity:
    """Evaluator and export agree on inlined operator semantics."""

    def test_string_plus_number_coercion(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, "5"),
            _make_node("S!A2", None, 3),
            _make_node("S!B1", "=S!A1+S!A2", None),
        )
        assert_codegen_matches_evaluator(graph, ["S!B1"])

    def test_blank_plus_number_treats_blank_as_zero(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, None),
            _make_node("S!A2", None, 3),
            _make_node("S!B1", "=S!A1+S!A2", None),
        )
        assert_codegen_matches_evaluator(graph, ["S!B1"])

    def test_date_string_arithmetic(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, "2024-01-01"),
            _make_node("S!A2", None, 1),
            _make_node("S!B1", "=S!A1+S!A2", None),
        )
        assert_codegen_matches_evaluator(graph, ["S!B1"])

    def test_mixed_type_comparison_uses_excel_ordering(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, 1),
            _make_node("S!A2", None, "10"),
            _make_node("S!B1", "=S!A1<S!A2", None),
            _make_node("S!B2", "=S!A1=S!A2", None),
        )
        assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2"])

    def test_concatenation_and_percent(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, 50),
            _make_node("S!A2", None, "pct"),
            _make_node("S!B1", "=S!A1%", None),
            _make_node("S!B2", '=S!A2&"!"', None),
        )
        assert_codegen_matches_evaluator(graph, ["S!B1", "S!B2"])

    def test_division_by_zero_raises_in_export(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=1/0", None))
        code = CodeGenerator(graph).generate(["S!A1"])
        assert "xl_add(" not in code
        assert "xl_number(" in code
        ns: dict[str, object] = {}
        exec(code, ns)
        compute_all = cast(Callable[[], object], ns["compute_all"])
        with pytest.raises(cast("type[BaseException]", ns["XlErrorException"])) as exc_info:
            compute_all()
        assert cast(Any, exc_info.value).code == XlError.DIV

    def test_invalid_power_raises_num(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=(-1)^0.5", None))
        assert_codegen_matches_evaluator(graph, ["S!A1"])

    def test_array_range_arithmetic_broadcast(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, 1),
            _make_node("S!A2", None, 2),
            _make_node("S!A3", None, 3),
            _make_node("S!B1", None, 10),
            _make_node("S!C1", "=SUM(OFFSET(S!A1,0,0,3,1)+S!B1)", None),
        )
        assert_codegen_matches_evaluator(graph, ["S!C1"])

    def test_generated_code_omits_removed_operator_wrappers(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, 1),
            _make_node("S!A2", None, 2),
            _make_node("S!B1", "=S!A1+S!A2", None),
            _make_node("S!B2", "=S!A1=S!A2", None),
            _make_node("S!B3", '=S!A1&"x"', None),
        )
        code = CodeGenerator(graph).generate(["S!B1", "S!B2", "S!B3"])
        for dead_wrapper in (
            "def xl_add(",
            "def xl_sub(",
            "def xl_mul(",
            "def xl_div(",
            "def xl_pow(",
            "def xl_eq(",
            "def xl_concat(",
            "def xl_neg(",
            "def xl_percent(",
        ):
            assert dead_wrapper not in code
        assert "xl_number(" in code
        assert "xl_compare(" in code


class TestArrayOperatorOperandBinding:
    """Guarded array operators bind operands once: no double-eval or blowup."""

    @staticmethod
    def _cell_body(code: str, key: str) -> str:
        out: list[str] = []
        grab = False
        for line in code.splitlines():
            if line.startswith(f"def {key}"):
                grab = True
            elif grab and line.startswith("def "):
                break
            if grab:
                out.append(line)
        return "\n".join(out)

    @classmethod
    def _array_graph(cls, formula: str) -> DependencyGraph:
        nodes = [_make_node("S!Z1", formula, None)]
        for col in "ABCD":
            for row in (1, 2, 3):
                nodes.append(_make_node(f"S!{col}{row}", None, float(row)))
        return _make_graph(*nodes)

    def test_guarded_operator_evaluates_each_operand_once(self) -> None:
        """A scalar-returning operand under a guard is emitted (evaluated) once."""
        graph = self._array_graph("=SUM(S!A1:S!A3)+S!B1")
        body = self._cell_body(CodeGenerator(graph).generate(["S!Z1"]), "cell_s_z1")
        assert body.count("xl_sum(") == 1
        assert body.count("xl_cell(ctx, 'S!B1')") == 1

    def test_nested_array_operators_do_not_duplicate_operands(self) -> None:
        """Each leaf operand appears once per textual use, not 3x per nesting level."""
        formula = "=((S!A1:A3+S!B1:B3)*(S!C1:C3-S!D1:D3))+((S!A1:A3-S!B1:B3)*(S!C1:C3+S!D1:D3))"
        body = self._cell_body(
            CodeGenerator(self._array_graph(formula)).generate(["S!Z1"]), "cell_s_z1"
        )
        # S!A1:A3 is used twice in the formula; without operand binding the
        # three-way guard duplicates it many times per nesting level.
        assert body.count("xl_range(ctx, 'S!A1:S!A3')") == 2

    def test_nested_array_operator_code_size_stays_linear(self) -> None:
        depth1 = self._cell_body(
            CodeGenerator(self._array_graph("=S!A1:A3+S!B1:B3")).generate(["S!Z1"]),
            "cell_s_z1",
        )
        depth3 = self._cell_body(
            CodeGenerator(
                self._array_graph(
                    "=((S!A1:A3+S!B1:B3)*(S!C1:C3-S!D1:D3))+((S!A1:A3-S!B1:B3)*(S!C1:C3+S!D1:D3))"
                )
            ).generate(["S!Z1"]),
            "cell_s_z1",
        )
        # Seven operators over eight leaves; the exponential guard produced ~27x.
        assert len(depth3) < 10 * len(depth1)

    def test_nested_array_operator_parity(self) -> None:
        graph = self._array_graph("=SUM((S!A1:A3+S!B1:B3)*(S!C1:C3-S!D1:D3))")
        assert_codegen_matches_evaluator(graph, ["S!Z1"])
