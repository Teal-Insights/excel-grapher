"""Tests for codegen module."""

from collections.abc import Callable
from typing import cast

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.parser import (
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
)
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.codegen import CodeGenerator, GraphLike, GraphNode
from tests.integration.utils.parity_harness import (
    CACHE_EVAL_SCAFFOLD_LINE_BUDGET,
    assert_cache_eval_scaffold_within_budget,
)


class _EmptyGraph:
    """Minimal GraphLike for AST emission tests that never touch nodes."""

    leaf_classification: dict[str, str] | None = None

    def get_node(self, address: str) -> GraphNode | None:
        return None

    def leaf_keys(self) -> list[str]:
        return []

    def formula_keys(self) -> list[str]:
        return []

    def get_dependencies(self, address: str) -> frozenset[str]:
        return frozenset()


def _ast_only_generator() -> CodeGenerator:
    graph: GraphLike = _EmptyGraph()
    return CodeGenerator(graph)


def _set_leaf_classification(graph: DependencyGraph, value: dict[str, str]) -> None:
    graph.leaf_classification = value


def _get_leaf_classification(graph: DependencyGraph) -> dict[str, str]:
    lc = graph.leaf_classification
    assert lc is not None
    return lc


class TestEmitAstLiterals:
    """Tests for _emit_ast with literal nodes."""

    @pytest.fixture
    def gen(self):
        """Create a CodeGenerator with a mock graph."""
        # For _emit_ast tests, we don't need a real graph
        return _ast_only_generator()

    def test_emit_number_integer(self, gen):
        """Integer numbers should emit without decimal."""
        assert gen._emit_ast(NumberNode(42.0)) == "42.0"

    def test_emit_number_float(self, gen):
        """Float numbers should preserve decimals."""
        assert gen._emit_ast(NumberNode(3.14)) == "3.14"

    def test_emit_number_negative(self, gen):
        """Negative numbers in NumberNode."""
        assert gen._emit_ast(NumberNode(-5.0)) == "-5.0"

    def test_emit_string_simple(self, gen):
        """Simple string literal."""
        assert gen._emit_ast(StringNode("hello")) == "'hello'"

    def test_emit_string_with_quotes(self, gen):
        """String containing double quotes uses single quote delimiters."""
        assert gen._emit_ast(StringNode('say "hi"')) == "'say \"hi\"'"

    def test_emit_string_empty(self, gen):
        """Empty string."""
        assert gen._emit_ast(StringNode("")) == "''"

    def test_emit_string_with_newline(self, gen):
        """String with newline should escape it."""
        assert gen._emit_ast(StringNode("line1\nline2")) == "'line1\\nline2'"

    def test_emit_bool_true(self, gen):
        """Boolean True."""
        assert gen._emit_ast(BoolNode(True)) == "True"

    def test_emit_bool_false(self, gen):
        """Boolean False."""
        assert gen._emit_ast(BoolNode(False)) == "False"

    def test_emit_error_na(self, gen):
        """#N/A error literals raise in the exported error channel."""
        assert gen._emit_ast(ErrorNode(XlError.NA)) == "xl_raise(XlError.NA)"

    def test_emit_error_value(self, gen):
        """#VALUE! error literals raise in the exported error channel."""
        assert gen._emit_ast(ErrorNode(XlError.VALUE)) == "xl_raise(XlError.VALUE)"

    def test_emit_error_ref(self, gen):
        """#REF! error literals raise in the exported error channel."""
        assert gen._emit_ast(ErrorNode(XlError.REF)) == "xl_raise(XlError.REF)"

    def test_emit_error_div(self, gen):
        """#DIV/0! error literals raise in the exported error channel."""
        assert gen._emit_ast(ErrorNode(XlError.DIV)) == "xl_raise(XlError.DIV)"


class TestEmitAstEmptyArg:
    """Tests for _emit_ast with EmptyArgNode (omitted function arguments)."""

    @pytest.fixture
    def gen(self):
        return _ast_only_generator()

    def test_emit_empty_arg(self, gen):
        """EmptyArgNode should emit None for omitted arguments."""
        assert gen._emit_ast(EmptyArgNode()) == "None"

    def test_emit_function_with_omitted_arg(self, gen):
        """Function call with omitted arguments should emit None in place."""
        node = FunctionCallNode(
            "INDEX",
            [
                RangeNode("Sheet1!A1", "Sheet1!B2"),
                EmptyArgNode(),
                NumberNode(1.0),
            ],
        )
        result = gen._emit_ast(node)
        assert "None" in result
        assert "1.0" in result


class TestEmitAstReferences:
    """Tests for _emit_ast with cell references and ranges."""

    @pytest.fixture
    def gen(self):
        return _ast_only_generator()

    def test_emit_cell_ref_simple(self, gen):
        """Simple cell reference."""
        assert gen._emit_ast(CellRefNode("Sheet1!A1")) == "xl_cell(ctx, 'Sheet1!A1')"

    def test_emit_cell_ref_quoted_sheet(self, gen):
        """Cell reference with quoted sheet name."""
        assert gen._emit_ast(CellRefNode("'My Sheet'!B2")) == "xl_cell(ctx, \"'My Sheet'!B2\")"

    def test_emit_range_1d_column(self, gen):
        """1D range (column) emits a lazy xl_range call."""
        result = gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!A3"))
        assert result == "xl_range(ctx, 'Sheet1!A1:A3')"

    def test_emit_range_1d_row(self, gen):
        """1D range (row) emits a lazy xl_range call."""
        result = gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!C1"))
        assert result == "xl_range(ctx, 'Sheet1!A1:C1')"

    def test_emit_range_2d(self, gen):
        """2D range emits a lazy xl_range call."""
        result = gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!B2"))
        assert result == "xl_range(ctx, 'Sheet1!A1:B2')"


class TestEmitAstOperators:
    """Tests for _emit_ast with operators."""

    @pytest.fixture
    def gen(self):
        return _ast_only_generator()

    def test_emit_binary_add(self, gen):
        """Addition inlines native + with xl_number coercion."""
        node = BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "(xl_number(1.0) + xl_number(2.0))"

    def test_emit_binary_subtract(self, gen):
        """Subtraction inlines native - with xl_number coercion."""
        node = BinaryOpNode("-", NumberNode(5.0), NumberNode(3.0))
        assert gen._emit_ast(node) == "(xl_number(5.0) - xl_number(3.0))"

    def test_emit_binary_multiply(self, gen):
        """Multiplication inlines native * with xl_number coercion."""
        node = BinaryOpNode("*", NumberNode(4.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "(xl_number(4.0) * xl_number(2.0))"

    def test_emit_binary_divide(self, gen):
        """Division inlines native / with Excel div-by-zero handling."""
        node = BinaryOpNode("/", NumberNode(10.0), NumberNode(2.0))
        assert gen._emit_ast(node) == (
            "((lambda _ln, _rn: (_ln / _rn if _rn != 0 else xl_raise(XlError.DIV)))"
            "(xl_number(10.0), xl_number(2.0)))"
        )

    def test_emit_binary_power(self, gen):
        """Exponentiation uses xl_pow_numbers on coerced operands."""
        node = BinaryOpNode("^", NumberNode(2.0), NumberNode(3.0))
        assert gen._emit_ast(node) == "xl_pow_numbers(xl_number(2.0), xl_number(3.0))"

    def test_emit_binary_concat(self, gen):
        """Concatenation inlines to_string + to_string."""
        node = BinaryOpNode("&", StringNode("a"), StringNode("b"))
        assert gen._emit_ast(node) == "(to_string('a') + to_string('b'))"

    def test_emit_binary_eq(self, gen):
        """Equality uses xl_compare for Excel semantics."""
        node = BinaryOpNode("=", NumberNode(1.0), NumberNode(1.0))
        assert gen._emit_ast(node) == "xl_compare('=', 1.0, 1.0)"

    def test_emit_binary_ne(self, gen):
        """Not equal uses xl_compare."""
        node = BinaryOpNode("<>", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_compare('<>', 1.0, 2.0)"

    def test_emit_binary_lt(self, gen):
        """Less than uses xl_compare."""
        node = BinaryOpNode("<", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_compare('<', 1.0, 2.0)"

    def test_emit_binary_gt(self, gen):
        """Greater than uses xl_compare."""
        node = BinaryOpNode(">", NumberNode(2.0), NumberNode(1.0))
        assert gen._emit_ast(node) == "xl_compare('>', 2.0, 1.0)"

    def test_emit_binary_le(self, gen):
        """Less than or equal uses xl_compare."""
        node = BinaryOpNode("<=", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_compare('<=', 1.0, 2.0)"

    def test_emit_binary_ge(self, gen):
        """Greater than or equal uses xl_compare."""
        node = BinaryOpNode(">=", NumberNode(2.0), NumberNode(1.0))
        assert gen._emit_ast(node) == "xl_compare('>=', 2.0, 1.0)"

    def test_emit_unary_minus(self, gen):
        """Unary minus inlines native negation with xl_number coercion."""
        node = UnaryOpNode("-", NumberNode(5.0))
        assert gen._emit_ast(node) == "(-xl_number(5.0))"

    def test_emit_nested_binary(self, gen):
        """Nested binary operations preserve inlined scalar paths."""
        inner = BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0))
        outer = BinaryOpNode("*", inner, NumberNode(3.0))
        inner_code = gen._emit_ast(inner)
        assert gen._emit_ast(outer) == f"(xl_number({inner_code}) * xl_number(3.0))"


class TestEmitAstFunctions:
    """Tests for _emit_ast with function calls."""

    @pytest.fixture
    def gen(self):
        return _ast_only_generator()

    def test_emit_function_no_args(self, gen):
        """Function with no arguments."""
        node = FunctionCallNode("TODAY", [])
        assert gen._emit_ast(node) == "xl_today()"

    def test_emit_na_raises_error_literal(self, gen):
        """NA() emits a raising error literal, not a sentinel-returning call."""
        node = FunctionCallNode("NA", [])
        assert gen._emit_ast(node) == "xl_raise(XlError.NA)"

    def test_emit_function_one_arg(self, gen):
        """Function with one argument."""
        node = FunctionCallNode("ABS", [NumberNode(-5.0)])
        assert gen._emit_ast(node) == "xl_abs(-5.0)"

    def test_emit_function_exp(self, gen):
        """EXP emits the shared runtime helper."""
        node = FunctionCallNode("EXP", [NumberNode(1.0)])
        assert gen._emit_ast(node) == "xl_exp(1.0)"

    def test_emit_function_multiple_args(self, gen):
        """Function with multiple arguments."""
        node = FunctionCallNode("SUM", [NumberNode(1.0), NumberNode(2.0), NumberNode(3.0)])
        assert gen._emit_ast(node) == "xl_sum(1.0, 2.0, 3.0)"

    def test_emit_function_nested(self, gen):
        """Nested function calls."""
        inner = FunctionCallNode("ABS", [NumberNode(-5.0)])
        outer = FunctionCallNode("SUM", [inner, NumberNode(10.0)])
        assert gen._emit_ast(outer) == "xl_sum(xl_abs(-5.0), 10.0)"

    def test_emit_function_with_cell_ref(self, gen):
        """Function with cell reference argument."""
        node = FunctionCallNode("SUM", [CellRefNode("Sheet1!A1"), CellRefNode("Sheet1!B1")])
        assert gen._emit_ast(node) == "xl_sum(xl_cell(ctx, 'Sheet1!A1'), xl_cell(ctx, 'Sheet1!B1'))"

    def test_emit_function_if(self, gen):
        """IF function - emits as Python conditional for lazy evaluation."""
        node = FunctionCallNode(
            "IF",
            [
                BinaryOpNode(">", CellRefNode("Sheet1!A1"), NumberNode(0.0)),
                StringNode("positive"),
                StringNode("non-positive"),
            ],
        )
        result = gen._emit_ast(node)
        # IF is emitted as a lazy conditional with raise-only boolean coercion.
        assert "xl_compare('>', xl_cell(ctx, 'Sheet1!A1'), 0.0)" in result
        assert "'positive'" in result
        assert "'non-positive'" in result
        assert "XlError" not in result
        assert "to_bool" not in result
        assert "xl_bool(" in result
        assert "if (_t1 :=" in result

    def test_emit_function_vlookup(self, gen):
        """VLOOKUP function."""
        node = FunctionCallNode(
            "VLOOKUP",
            [
                CellRefNode("Sheet1!A1"),
                RangeNode("Sheet1!B1", "Sheet1!C10"),
                NumberNode(2.0),
                BoolNode(False),
            ],
        )
        result = gen._emit_ast(node)
        assert "xl_vlookup(" in result
        assert "xl_cell(ctx, 'Sheet1!A1')" in result


# --- Helper functions for graph creation ---


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


def _build_codegen_diamond_graph(*, reverse_insertion: bool) -> DependencyGraph:
    nodes = [
        _make_node("Sheet1!A1", None, 1.0),
        _make_node("Sheet1!B1", "=Sheet1!A1+1", None),
        _make_node("Sheet1!C1", "=Sheet1!A1*2", None),
        _make_node("Sheet1!D1", "=Sheet1!B1+Sheet1!C1", None),
    ]
    if reverse_insertion:
        nodes = list(reversed(nodes))

    graph = DependencyGraph(sheet_order=["Sheet1"])
    for node in nodes:
        graph.add_node(node)
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!C1", "Sheet1!A1")
    graph.add_edge("Sheet1!D1", "Sheet1!B1")
    graph.add_edge("Sheet1!D1", "Sheet1!C1")
    return graph


class TestDeterministicCodegenOrdering:
    def test_generate_modules_emits_formula_functions_in_workbook_order(self) -> None:
        graph = _build_codegen_diamond_graph(reverse_insertion=True)
        files = CodeGenerator(graph).generate_modules(["Sheet1!D1"])
        internals = files["internals.py"]

        positions = [internals.index(f"def cell_sheet1_{col}1(") for col in ("b", "c", "d")]
        assert positions == sorted(positions)

    def test_generate_modules_formula_order_is_independent_of_node_insertion_order(self) -> None:
        forward = CodeGenerator(
            _build_codegen_diamond_graph(reverse_insertion=False)
        ).generate_modules(["Sheet1!D1"])
        reverse = CodeGenerator(
            _build_codegen_diamond_graph(reverse_insertion=True)
        ).generate_modules(["Sheet1!D1"])
        assert forward["internals.py"] == reverse["internals.py"]


class TestEmitCell:
    """Tests for _emit_cell method."""

    def test_emit_leaf_cell(self):
        """Leaf cells are exported as data, not functions."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        with pytest.raises(ValueError):
            _ = gen._emit_cell("Sheet1!A1")

    def test_emit_leaf_cell_string(self):
        """Leaf cells are exported as data, not functions."""
        graph = _make_graph(_make_node("Sheet1!A1", None, "hello"))
        gen = CodeGenerator(graph)
        with pytest.raises(ValueError):
            _ = gen._emit_cell("Sheet1!A1")

    def test_emit_leaf_cell_none(self):
        """Leaf cells are exported as data, not functions."""
        graph = _make_graph(_make_node("Sheet1!A1", None, None))
        gen = CodeGenerator(graph)
        with pytest.raises(ValueError):
            _ = gen._emit_cell("Sheet1!A1")

    def test_emit_formula_cell(self):
        """Formula cell should emit parsed formula."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 100.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
        )
        gen = CodeGenerator(graph)
        code = gen._emit_cell("Sheet1!B1")
        assert "def cell_sheet1_b1(ctx):" in code
        assert "xl_cell(ctx, 'Sheet1!A1')" in code
        assert "xl_number(" in code
        assert "xl_mul(" not in code

    def test_emit_formula_cell_with_function(self):
        """Formula cell with function call."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!A2", None, 2.0),
            _make_node("Sheet1!B1", "=SUM(Sheet1!A1, Sheet1!A2)", None),
        )
        gen = CodeGenerator(graph)
        code = gen._emit_cell("Sheet1!B1")
        assert "xl_sum(" in code
        assert "xl_cell(ctx, 'Sheet1!A1')" in code
        assert "xl_cell(ctx, 'Sheet1!A2')" in code

    def test_emit_cell_includes_docstring(self):
        """Emitted cell function should include docstring with original formula."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 100.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
        )
        gen = CodeGenerator(graph)
        code = gen._emit_cell("Sheet1!B1")
        # Should have a docstring
        assert '"""' in code or "'''" in code

    def test_emit_cell_docstring_escapes_quotes(self):
        """Docstring should stay valid when formulas contain quotes."""
        graph = _make_graph(
            _make_node("Sheet1!B1", '=Sheet1!J5&" Swap rate"', None),
        )
        gen = CodeGenerator(graph)
        code = gen._emit_cell("Sheet1!B1")
        exec(code, {})


class TestGenerate:
    """Tests for generate() method."""

    def test_generate_defaults_to_graph_targets(self):
        """generate() should use graph target metadata when targets are omitted."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!A1+1", None),
        )
        b1 = graph._get_internal_node("Sheet1!B1")
        c1 = graph._get_internal_node("Sheet1!C1")
        assert b1 is not None
        assert c1 is not None
        b1.is_target = True
        c1.is_target = False
        graph.add_edge("Sheet1!B1", "Sheet1!A1")
        graph.add_edge("Sheet1!C1", "Sheet1!A1")

        code = CodeGenerator(graph).generate()

        assert "TARGETS = {" in code
        assert "    'Sheet1!B1': xl_cell," in code
        assert "    'Sheet1!C1': xl_cell," not in code

    def test_generate_without_targets_raises_when_graph_has_no_targets(self):
        """generate() without explicit targets should fail on targetless graphs."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))

        with pytest.raises(ValueError, match="No export targets were provided"):
            _ = CodeGenerator(graph).generate()

    def test_generate_caches_parsed_asts(self, monkeypatch):
        """generate() should not repeatedly parse the same cell formulas."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, 20.0),
            _make_node("Sheet1!B1", "=Sheet1!A1+Sheet1!A2", None),
            _make_node("Sheet1!C1", "=Sheet1!B1*2", None),
        )
        gen = CodeGenerator(graph)

        # Monkeypatch exporter.codegen.parse; implementation lives there (evaluator.codegen is a shim).
        import excel_grapher.exporter.codegen as codegen_module

        original_parse = codegen_module.parse
        calls: list[str] = []

        def counting_parse(formula: str):
            calls.append(formula)
            return original_parse(formula)

        monkeypatch.setattr(codegen_module, "parse", counting_parse)

        _ = gen.generate(["Sheet1!C1"])

        # Only formula cells should be parsed, and each should be parsed once.
        assert len(calls) == 2


class TestCodeGeneratorContextManager:
    def test_context_manager_clears_transient_codegen_state(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
        )
        graph.add_edge("Sheet1!B1", "Sheet1!A1")
        gen = CodeGenerator(graph)

        with gen as scoped:
            _ = scoped.generate(["Sheet1!B1"])
            assert scoped._ast_cache
            assert scoped._emitted

        assert gen._ast_cache == {}
        assert gen._emitted == set()
        assert gen._formula_cell_address is None
        assert gen._temp_var_counter == 0

    def test_generate_includes_imports(self):
        """Generated code should include necessary imports."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])
        assert "class EvalContext" in code
        assert "def xl_cell(" in code
        # Should be standalone - no excel_evaluator imports
        assert "from excel_evaluator" not in code

    def test_emitted_cache_eval_scaffold_within_line_budget(self) -> None:
        """Shared _evaluate_address helper keeps export scaffold under line budget."""
        graph = _make_graph(_make_node("Sheet1!A1", "=1+1", None))
        code = CodeGenerator(graph).generate(["Sheet1!A1"])
        line_count = assert_cache_eval_scaffold_within_budget(code)
        assert line_count <= CACHE_EVAL_SCAFFOLD_LINE_BUDGET

    def test_generate_runtime_imports_do_not_redefine_callable(self):
        """Generated runtime should not import Callable twice (ruff F811)."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])

        # Callable should not be imported from typing at all.
        assert "from typing import Any, Callable" not in code
        assert "from typing import Callable" not in code

        # Callable should be imported from collections.abc only once in the flattened runtime.
        assert code.count("Callable") >= 1  # sanity: used in runtime annotations
        assert code.count("from collections.abc import Callable") <= 1
        assert code.count("from collections.abc import Callable,") <= 1

    def test_generate_includes_all_dependencies(self):
        """Generated code should include all dependent cells."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!B1+Sheet1!A1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1"])
        # Leaf inputs are data; formulas are functions.
        assert "DEFAULT_INPUTS" in code
        assert "    'Sheet1!A1': 10.0," in code
        assert "def cell_sheet1_b1(ctx):" in code
        assert "def cell_sheet1_c1(ctx):" in code

    def test_generate_splits_constants_by_type(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, "hi"),
            _make_node("Sheet1!A3", None, None),
            _make_node("Sheet1!A4", None, True),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!A1", "Sheet1!A2", "Sheet1!A3", "Sheet1!A4"],
            constant_types={"number", "string"},
        )
        assert "CONSTANTS = {" in code
        assert "    'Sheet1!A1': 10.0," in code
        assert "    'Sheet1!A2': 'hi'," in code
        assert "    'Sheet1!A3': 0," in code
        assert "DEFAULT_INPUTS = {" in code
        assert "    'Sheet1!A4': True," in code
        assert code.index("DEFAULT_INPUTS = {") < code.index("CONSTANTS = {")

    def test_generate_constant_ranges_override_types(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, "hi"),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!A1", "Sheet1!A2"],
            constant_types={"string"},
            constant_ranges=["Sheet1!A1"],
        )
        assert "CONSTANTS = {" in code
        assert "    'Sheet1!A1': 10.0," in code
        assert "    'Sheet1!A2': 'hi'," in code
        assert "DEFAULT_INPUTS" in code

    def test_generate_input_ranges_override_constant_types(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, 20.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!A1", "Sheet1!A2"],
            constant_types={"number"},
            input_ranges=["Sheet1!A1"],
        )
        assert "CONSTANTS = {" in code
        assert "    'Sheet1!A2': 20.0," in code
        assert "    'Sheet1!A1': 10.0," in code
        assert code.index("DEFAULT_INPUTS = {") < code.index("CONSTANTS = {")

    def test_generate_input_ranges_override_constant_ranges(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!A2", None, 2.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!A1", "Sheet1!A2"],
            constant_ranges=["Sheet1!A1:A2"],
            input_ranges=["Sheet1!A1"],
        )
        assert "    'Sheet1!A1': 1.0," in code
        assert "    'Sheet1!A2': 2.0," in code
        assert code.index("DEFAULT_INPUTS = {") < code.index("CONSTANTS = {")

    def test_generate_input_ranges_override_graph_leaf_classification(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, 20.0),
        )
        _set_leaf_classification(graph, {"Sheet1!A1": "constant", "Sheet1!A2": "constant"})
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!A1", "Sheet1!A2"],
            input_ranges=["Sheet1!A1"],
        )
        assert "    'Sheet1!A1': 10.0," in code
        assert "    'Sheet1!A2': 20.0," in code
        assert code.index("DEFAULT_INPUTS = {") < code.index("CONSTANTS = {")

    def test_classify_leaf_nodes_input_ranges(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!A2", None, 2.0),
        )
        gen = CodeGenerator(graph)
        inputs, constants = gen.classify_leaf_nodes(
            ["Sheet1!A1", "Sheet1!A2"],
            constant_types={"number"},
            input_ranges=["Sheet1!A1"],
        )
        assert inputs == {"Sheet1!A1"}
        assert constants == {"Sheet1!A2"}

    def test_generate_constant_blanks(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, None),
            _make_node("Sheet1!A2", None, 3.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!A1", "Sheet1!A2"],
            constant_blanks=True,
        )
        assert "CONSTANTS = {" in code
        assert "    'Sheet1!A1': 0," in code
        assert "DEFAULT_INPUTS = {" in code
        assert "    'Sheet1!A2': 3.0," in code

    def test_generate_uses_graph_leaf_classification_by_default(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, 20.0),
        )
        _set_leaf_classification(graph, {"Sheet1!A1": "constant", "Sheet1!A2": "input"})
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1", "Sheet1!A2"])
        assert "CONSTANTS = {" in code
        assert "    'Sheet1!A1': 10.0," in code
        assert "DEFAULT_INPUTS = {" in code
        assert "    'Sheet1!A2': 20.0," in code

    def test_generate_kwargs_override_graph_leaf_classification(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, 20.0),
        )
        _set_leaf_classification(graph, {"Sheet1!A1": "input", "Sheet1!A2": "input"})
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!A1", "Sheet1!A2"],
            constant_types={"number"},
        )
        assert "CONSTANTS = {" in code
        assert "    'Sheet1!A1': 10.0," in code
        assert "    'Sheet1!A2': 20.0," in code
        assert "DEFAULT_INPUTS" in code

    def test_classify_leaf_nodes_attaches_to_graph(self):
        graph = _make_graph(
            _make_node("Sheet1!A1", None, None),
            _make_node("Sheet1!A2", None, 4.0),
        )
        gen = CodeGenerator(graph)
        inputs, constants = gen.classify_leaf_nodes(
            ["Sheet1!A1", "Sheet1!A2"],
            constant_blanks=True,
            attach_to_graph=True,
        )
        assert inputs == {"Sheet1!A2"}
        assert constants == {"Sheet1!A1"}
        assert _get_leaf_classification(graph) == {
            "Sheet1!A1": "constant",
            "Sheet1!A2": "input",
        }

    def test_generate_includes_entry_point(self):
        """Generated code should include compute_all entry point."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])
        assert (
            "def compute_all(ctx: EvalContext | None = None, *, "
            "inputs: dict[str, object] | None = None) -> dict[str, object]:"
        ) in code
        assert "'Sheet1!A1'" in code

    def test_generate_includes_empty_series_binding_discovery_helpers(self):
        """Generated code should expose discovery helpers even without bindings."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        code = CodeGenerator(graph).generate(["Sheet1!A1"])

        namespace: dict[str, object] = {}
        exec(code, namespace)
        list_setters = cast(Callable[[], list[str]], namespace["list_setters"])
        list_readers = cast(Callable[[], list[str]], namespace["list_readers"])
        list_computes = cast(Callable[[], list[str]], namespace["list_computes"])

        assert list_setters() == []
        assert list_readers() == []
        assert list_computes() == []

    def test_generate_entrypoint_uses_target_map(self):
        """Generated compute_all should iterate a shared targets map."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])
        assert "TARGETS = {" in code
        assert (
            "    return {target: handler(ctx, target) for target, handler in TARGETS.items()}"
            in code
        )

    def test_generate_entrypoint_emits_ranges_for_contiguous_row(self):
        graph = _make_graph(
            _make_node("Sheet1!C1", None, 1.0),
            _make_node("Sheet1!D1", None, 2.0),
            _make_node("Sheet1!E1", None, 3.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1", "Sheet1!D1", "Sheet1!E1"])
        assert "Sheet1!C1:E1" in code
        assert "'Sheet1!C1:E1': xl_range_rows" in code

    def test_generate_deduplication(self):
        """Cells should only be emitted once even if referenced multiple times."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!A1+Sheet1!B1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1"])
        # A1 is referenced by both B1 and C1, but should only have one DEFAULT_INPUTS entry
        assert code.count("    'Sheet1!A1':") == 1


class TestGenerateNamedRanges:
    """Direct `targets` accept the same forms as graph build targets."""

    @staticmethod
    def _graph_with_named_ranges(*nodes: Node) -> DependencyGraph:
        graph = _make_graph(*nodes)
        graph.sheet_order = ["Sheet1"]
        graph.named_ranges = {"OneCell": ("Sheet1", "C1")}
        graph.named_range_ranges = {"BeeCol": ("Sheet1", "B1", "B3")}
        return graph

    def test_generate_expands_defined_name_single_cell(self):
        graph = self._graph_with_named_ranges(
            _make_node("Sheet1!B1", None, 1.0),
            _make_node("Sheet1!C1", "=Sheet1!B1+1", None),
        )
        code = CodeGenerator(graph).generate(["OneCell"])
        assert (
            "def compute_all(ctx: EvalContext | None = None, *, "
            "inputs: dict[str, object] | None = None) -> dict[str, object]:"
        ) in code
        assert "'Sheet1!C1': xl_cell" in code

    def test_generate_expands_defined_name_range(self):
        graph = self._graph_with_named_ranges(
            _make_node("Sheet1!B1", None, 1.0),
            _make_node("Sheet1!B2", None, 2.0),
            _make_node("Sheet1!B3", None, 3.0),
        )
        code = CodeGenerator(graph).generate(["BeeCol"])
        assert "'Sheet1!B1:B3': xl_range_rows" in code

    def test_generate_modules_expands_defined_name(self):
        graph = self._graph_with_named_ranges(
            _make_node("Sheet1!B1", None, 1.0),
            _make_node("Sheet1!B2", None, 2.0),
            _make_node("Sheet1!B3", None, 3.0),
        )
        files = CodeGenerator(graph).generate_modules(["BeeCol"])
        api_py = files["api.py"]
        assert "'Sheet1!B1:B3': xl_range_rows" in api_py

    def test_generate_unknown_defined_name_raises(self):
        graph = self._graph_with_named_ranges(_make_node("Sheet1!A1", None, 1.0))
        gen = CodeGenerator(graph)
        with pytest.raises(ValueError, match="NoSuchName"):
            _ = gen.generate(["NoSuchName"])


class TestGeneratedCodeExecution:
    def test_generated_code_executes(self):
        """Generated code should be executable and produce correct results."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!A1+Sheet1!B1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1"])

        # Execute the generated code
        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()

        assert result["Sheet1!C1"] == 30.0  # 10 + 20

    def test_generated_code_allows_overriding_inputs(self):
        """Callers can override exported-time leaf values via compute_all(inputs=...)."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!B1"])
        namespace: dict = {}
        exec(code, namespace)
        compute_all = namespace["compute_all"]

        assert compute_all()["Sheet1!B1"] == 20.0
        assert compute_all(inputs={"Sheet1!A1": 7.0})["Sheet1!B1"] == 14.0

    def test_generated_code_caches_formula_results_per_run(self):
        """Generated code should compute formula cells only once per ctx."""
        graph = _make_graph(
            _make_node("Sheet1!A1", "=1+1", None),
            _make_node("Sheet1!C1", "=Sheet1!A1+Sheet1!A1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1"])

        namespace: dict = {}
        exec(code, namespace)
        eval_context = namespace["EvalContext"]
        xl_cell = namespace["xl_cell"]
        resolver = namespace["_resolve_formula"]
        base_inputs = dict(namespace["DEFAULT_INPUTS"])

        call_count = {"A1": 0}
        original = namespace["cell_sheet1_a1"]

        def wrapped(ctx):
            call_count["A1"] += 1
            return original(ctx)

        namespace["cell_sheet1_a1"] = wrapped

        ctx = eval_context(inputs=dict(base_inputs), resolver=resolver)
        xl_cell(ctx, "Sheet1!C1")
        assert call_count["A1"] == 1

        xl_cell(ctx, "Sheet1!C1")
        assert call_count["A1"] == 1

        ctx2 = eval_context(inputs=dict(base_inputs), resolver=resolver)
        xl_cell(ctx2, "Sheet1!C1")
        assert call_count["A1"] == 2

    def test_generated_code_make_context_reuses_cache_across_entrypoints(self):
        """Reusing a ctx across compute_all calls should preserve cached results."""
        graph = _make_graph(
            _make_node("Sheet1!A1", "=1+1", None),
            _make_node("Sheet1!C1", "=Sheet1!A1+Sheet1!A1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1"])

        namespace: dict = {}
        exec(code, namespace)
        make_context = namespace["make_context"]
        compute_all = namespace["compute_all"]

        call_count = {"A1": 0}
        original = namespace["cell_sheet1_a1"]

        def wrapped(ctx):
            call_count["A1"] += 1
            return original(ctx)

        namespace["cell_sheet1_a1"] = wrapped

        ctx = make_context()
        _ = compute_all(ctx=ctx)
        assert call_count["A1"] == 1

        _ = compute_all(ctx=ctx)
        assert call_count["A1"] == 1

    def test_generated_code_prefers_ctx_over_inputs_with_warning(self):
        """compute_all should warn when both ctx and inputs are provided."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])

        namespace: dict = {}
        exec(code, namespace)
        make_context = namespace["make_context"]
        compute_all = namespace["compute_all"]

        ctx = make_context()
        with pytest.warns(UserWarning, match="inputs will be ignored"):
            _ = compute_all(ctx=ctx, inputs={"Sheet1!A1": 99.0})

    def test_generated_code_partial_cache_invalidation_on_input_change(self):
        """Slim exports recompute via fresh context; full exports invalidate selectively."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!B1+1", None),
            _make_node("Sheet1!D1", None, 3.0),
            _make_node("Sheet1!E1", "=Sheet1!D1+1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1", "Sheet1!E1"])

        namespace: dict = {}
        exec(code, namespace)
        make_context = namespace["make_context"]
        compute_all = namespace["compute_all"]

        assert "def set_inputs(" not in code

        call_count = {"C1": 0, "E1": 0}
        original_c1 = namespace["cell_sheet1_c1"]
        original_e1 = namespace["cell_sheet1_e1"]

        def wrapped_c1(ctx):
            call_count["C1"] += 1
            return original_c1(ctx)

        def wrapped_e1(ctx):
            call_count["E1"] += 1
            return original_e1(ctx)

        namespace["cell_sheet1_c1"] = wrapped_c1
        namespace["cell_sheet1_e1"] = wrapped_e1

        ctx = make_context()
        result = compute_all(ctx=ctx)
        assert result["Sheet1!C1"] == 21.0
        assert result["Sheet1!E1"] == 4.0
        assert call_count == {"C1": 1, "E1": 1}

        result = compute_all(inputs={"Sheet1!A1": 7.0})
        assert result["Sheet1!C1"] == 15.0
        assert result["Sheet1!E1"] == 4.0
        assert call_count == {"C1": 2, "E1": 2}

    def test_generated_code_with_sum(self):
        """Generated code with SUM function should execute correctly."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 1.0),
            _make_node("Sheet1!A2", None, 2.0),
            _make_node("Sheet1!A3", None, 3.0),
            _make_node("Sheet1!B1", "=SUM(Sheet1!A1, Sheet1!A2, Sheet1!A3)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!B1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()

        assert result["Sheet1!B1"] == 6.0

    def test_generate_multiple_targets(self):
        """Generate code for multiple target cells."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!A1*3", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!B1", "Sheet1!C1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()

        assert result["Sheet1!B1:C1"] == [[20.0, 30.0]]


class TestIndexPrunedRangeCodegen:
    """INDEX over a range must not read graph-pruned cells (issue #201)."""

    def test_index_range_codegen_skips_pruned_cells_and_runs(self) -> None:
        graph = _make_graph(
            _make_node("Inputs!A2", None, "Borvelia"),
            _make_node("Inputs!B2", None, 60),
            _make_node("Inputs!C2", None, "stylized emerging market"),
            _make_node("Inputs!E1", None, "Borvelia"),
            _make_node(
                "Inputs!F1",
                "=INDEX(Inputs!A2:Inputs!C2, MATCH(Inputs!E1, Inputs!A2:Inputs!A2, 0), 2)",
                None,
            ),
        )
        graph.add_edge("Inputs!F1", "Inputs!A2")
        graph.add_edge("Inputs!F1", "Inputs!B2")
        graph.add_edge("Inputs!F1", "Inputs!E1")

        gen = CodeGenerator(graph)
        code = gen.generate(["Inputs!F1"])

        assert "xl_cell(ctx, 'Inputs!C2')" not in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert "Inputs!F1" in result
