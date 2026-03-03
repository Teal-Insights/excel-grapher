"""Tests for codegen module."""

import pytest

from excel_grapher.evaluator.codegen import CodeGenerator
from excel_grapher.evaluator.parser import (
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
)
from excel_grapher.evaluator.types import XlError
from excel_grapher.evaluator.name_utils import parse_address
from excel_grapher import DependencyGraph
from excel_grapher import Node


class TestEmitAstLiterals:
    """Tests for _emit_ast with literal nodes."""

    @pytest.fixture
    def gen(self):
        """Create a CodeGenerator with a mock graph."""
        # For _emit_ast tests, we don't need a real graph
        return CodeGenerator(None)  # type: ignore

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
        assert gen._emit_ast(StringNode('say "hi"')) == '\'say "hi"\''

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
        """#N/A error."""
        assert gen._emit_ast(ErrorNode(XlError.NA)) == "XlError.NA"

    def test_emit_error_value(self, gen):
        """#VALUE! error."""
        assert gen._emit_ast(ErrorNode(XlError.VALUE)) == "XlError.VALUE"

    def test_emit_error_ref(self, gen):
        """#REF! error."""
        assert gen._emit_ast(ErrorNode(XlError.REF)) == "XlError.REF"

    def test_emit_error_div(self, gen):
        """#DIV/0! error."""
        assert gen._emit_ast(ErrorNode(XlError.DIV)) == "XlError.DIV"


class TestEmitAstReferences:
    """Tests for _emit_ast with cell references and ranges."""

    @pytest.fixture
    def gen(self):
        return CodeGenerator(None)  # type: ignore

    def test_emit_cell_ref_simple(self, gen):
        """Simple cell reference."""
        assert gen._emit_ast(CellRefNode("Sheet1!A1")) == "xl_cell(ctx, 'Sheet1!A1')"

    def test_emit_cell_ref_quoted_sheet(self, gen):
        """Cell reference with quoted sheet name."""
        assert (
            gen._emit_ast(CellRefNode("'My Sheet'!B2")) == "xl_cell(ctx, \"'My Sheet'!B2\")"
        )

    def test_emit_range_1d_column(self, gen):
        """1D range (column)."""
        result = gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!A3"))
        # Should expand to list of cell calls
        assert "xl_cell(ctx, 'Sheet1!A1')" in result
        assert "xl_cell(ctx, 'Sheet1!A2')" in result
        assert "xl_cell(ctx, 'Sheet1!A3')" in result
        # Should be an object-dtype ndarray
        assert result.startswith("np.array([")
        assert result.endswith("dtype=object)")

    def test_emit_range_1d_row(self, gen):
        """1D range (row)."""
        result = gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!C1"))
        assert "xl_cell(ctx, 'Sheet1!A1')" in result
        assert "xl_cell(ctx, 'Sheet1!B1')" in result
        assert "xl_cell(ctx, 'Sheet1!C1')" in result

    def test_emit_range_2d(self, gen):
        """2D range."""
        result = gen._emit_ast(RangeNode("Sheet1!A1", "Sheet1!B2"))
        assert "xl_cell(ctx, 'Sheet1!A1')" in result
        assert "xl_cell(ctx, 'Sheet1!B1')" in result
        assert "xl_cell(ctx, 'Sheet1!A2')" in result
        assert "xl_cell(ctx, 'Sheet1!B2')" in result


class TestEmitAstOperators:
    """Tests for _emit_ast with operators."""

    @pytest.fixture
    def gen(self):
        return CodeGenerator(None)  # type: ignore

    def test_emit_binary_add(self, gen):
        """Addition operator uses xl_add for error propagation."""
        node = BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_add(1.0, 2.0)"

    def test_emit_binary_subtract(self, gen):
        """Subtraction operator uses xl_sub for error propagation."""
        node = BinaryOpNode("-", NumberNode(5.0), NumberNode(3.0))
        assert gen._emit_ast(node) == "xl_sub(5.0, 3.0)"

    def test_emit_binary_multiply(self, gen):
        """Multiplication operator uses xl_mul for error propagation."""
        node = BinaryOpNode("*", NumberNode(4.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_mul(4.0, 2.0)"

    def test_emit_binary_divide(self, gen):
        """Division operator uses xl_div for safe division."""
        node = BinaryOpNode("/", NumberNode(10.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_div(10.0, 2.0)"

    def test_emit_binary_power(self, gen):
        """Exponentiation operator uses xl_pow for error propagation."""
        node = BinaryOpNode("^", NumberNode(2.0), NumberNode(3.0))
        assert gen._emit_ast(node) == "xl_pow(2.0, 3.0)"

    def test_emit_binary_concat(self, gen):
        """Concatenation operator (& -> xl_concat)."""
        node = BinaryOpNode("&", StringNode("a"), StringNode("b"))
        assert gen._emit_ast(node) == "xl_concat('a', 'b')"

    def test_emit_binary_eq(self, gen):
        """Equality operator (= -> xl_eq for Excel semantics)."""
        node = BinaryOpNode("=", NumberNode(1.0), NumberNode(1.0))
        assert gen._emit_ast(node) == "xl_eq(1.0, 1.0)"

    def test_emit_binary_ne(self, gen):
        """Not equal operator (<> -> xl_ne)."""
        node = BinaryOpNode("<>", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_ne(1.0, 2.0)"

    def test_emit_binary_lt(self, gen):
        """Less than operator (< -> xl_lt)."""
        node = BinaryOpNode("<", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_lt(1.0, 2.0)"

    def test_emit_binary_gt(self, gen):
        """Greater than operator (> -> xl_gt)."""
        node = BinaryOpNode(">", NumberNode(2.0), NumberNode(1.0))
        assert gen._emit_ast(node) == "xl_gt(2.0, 1.0)"

    def test_emit_binary_le(self, gen):
        """Less than or equal operator (<= -> xl_le)."""
        node = BinaryOpNode("<=", NumberNode(1.0), NumberNode(2.0))
        assert gen._emit_ast(node) == "xl_le(1.0, 2.0)"

    def test_emit_binary_ge(self, gen):
        """Greater than or equal operator (>= -> xl_ge)."""
        node = BinaryOpNode(">=", NumberNode(2.0), NumberNode(1.0))
        assert gen._emit_ast(node) == "xl_ge(2.0, 1.0)"

    def test_emit_unary_minus(self, gen):
        """Unary minus uses xl_neg for error propagation."""
        node = UnaryOpNode("-", NumberNode(5.0))
        assert gen._emit_ast(node) == "xl_neg(5.0)"

    def test_emit_nested_binary(self, gen):
        """Nested binary operations."""
        # (1 + 2) * 3
        inner = BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0))
        outer = BinaryOpNode("*", inner, NumberNode(3.0))
        assert gen._emit_ast(outer) == "xl_mul(xl_add(1.0, 2.0), 3.0)"


class TestEmitAstFunctions:
    """Tests for _emit_ast with function calls."""

    @pytest.fixture
    def gen(self):
        return CodeGenerator(None)  # type: ignore

    def test_emit_function_no_args(self, gen):
        """Function with no arguments."""
        node = FunctionCallNode("NA", [])
        assert gen._emit_ast(node) == "xl_na()"

    def test_emit_function_one_arg(self, gen):
        """Function with one argument."""
        node = FunctionCallNode("ABS", [NumberNode(-5.0)])
        assert gen._emit_ast(node) == "xl_abs(-5.0)"

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
        # IF is emitted as lazy conditional: error check, then conditional expression
        assert "xl_gt(xl_cell(ctx, 'Sheet1!A1'), 0.0)" in result
        assert "'positive'" in result
        assert "'non-positive'" in result
        assert "XlError" in result  # Error propagation check
        # Condition is evaluated once, coerced via to_bool(), then used in a lazy conditional.
        assert "to_bool" in result
        assert "if _t2 else" in result

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
        assert "xl_mul(" in code  # Uses wrapper for error propagation

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

    def test_generate_caches_parsed_asts(self, monkeypatch):
        """generate() should not repeatedly parse the same cell formulas."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!A2", None, 20.0),
            _make_node("Sheet1!B1", "=Sheet1!A1+Sheet1!A2", None),
            _make_node("Sheet1!C1", "=Sheet1!B1*2", None),
        )
        gen = CodeGenerator(graph)

        # Monkeypatch codegen.parse (not parser.parse) because codegen imports parse directly.
        import excel_grapher.evaluator.codegen as codegen_module

        original_parse = codegen_module.parse
        calls: list[str] = []

        def counting_parse(formula: str):
            calls.append(formula)
            return original_parse(formula)

        monkeypatch.setattr(codegen_module, "parse", counting_parse)

        _ = gen.generate(["Sheet1!C1"])

        # Only formula cells should be parsed, and each should be parsed once.
        assert len(calls) == 2

    def test_generate_includes_imports(self):
        """Generated code should include necessary imports."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])
        assert "class EvalContext" in code
        assert "def xl_cell(" in code
        # Should be standalone - no excel_evaluator imports
        assert "from excel_evaluator" not in code

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
        setattr(
            graph,
            "leaf_classification",
            {"Sheet1!A1": "constant", "Sheet1!A2": "input"},
        )
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
        setattr(
            graph,
            "leaf_classification",
            {"Sheet1!A1": "input", "Sheet1!A2": "input"},
        )
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
        assert getattr(graph, "leaf_classification") == {
            "Sheet1!A1": "constant",
            "Sheet1!A2": "input",
        }

    def test_generate_includes_entry_point(self):
        """Generated code should include compute_all entry point."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])
        assert "def compute_all(inputs=None, *, ctx=None):" in code
        assert "'Sheet1!A1'" in code

    def test_generate_entrypoint_uses_target_map(self):
        """Generated compute_all should iterate a shared targets map."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 100.0))
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!A1"])
        assert "TARGETS = {" in code
        assert "    return {target: handler(ctx, target) for target, handler in TARGETS.items()}" in code

    def test_generate_entrypoint_emits_ranges_for_contiguous_row(self):
        graph = _make_graph(
            _make_node("Sheet1!C1", None, 1.0),
            _make_node("Sheet1!D1", None, 2.0),
            _make_node("Sheet1!E1", None, 3.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["Sheet1!C1", "Sheet1!D1", "Sheet1!E1"])
        assert "Sheet1!C1:Sheet1!E1" in code
        assert "'Sheet1!C1:Sheet1!E1': xl_range" in code

    def test_generate_entrypoints_emit_named_functions(self):
        """Generated code should include compute_* entrypoints for each mapping key."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!B1+1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!B1"],
            entrypoints={
                "outputs-a": ["Sheet1!B1"],
                "outputs_b": ["Sheet1!C1"],
            },
        )
        assert "def compute_outputs_a(inputs=None, *, ctx=None):" in code
        assert "def compute_outputs_b(inputs=None, *, ctx=None):" in code

    def test_generated_code_entrypoints_share_ctx_cache(self):
        """Entrypoints should share cached dependencies when ctx is reused."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!B1+1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!B1"],
            entrypoints={"first": ["Sheet1!B1"], "second": ["Sheet1!C1"]},
        )

        namespace: dict = {}
        exec(code, namespace)
        make_context = namespace["make_context"]
        compute_first = namespace["compute_first"]
        compute_second = namespace["compute_second"]
        call_count = {"B1": 0}
        original_b1 = namespace["cell_sheet1_b1"]

        def wrapped(ctx):
            call_count["B1"] += 1
            return original_b1(ctx)

        namespace["cell_sheet1_b1"] = wrapped

        ctx = make_context()
        _ = compute_first(ctx=ctx)
        _ = compute_second(ctx=ctx)
        assert call_count["B1"] == 1

    def test_generate_entrypoints_rejects_name_collision(self):
        """Entrypoints with colliding normalized names should error."""
        graph = _make_graph(_make_node("Sheet1!A1", None, 10.0))
        gen = CodeGenerator(graph)
        with pytest.raises(ValueError, match="normalize to the same identifier"):
            _ = gen.generate(
                ["Sheet1!A1"],
                entrypoints={"outputs-a": ["Sheet1!A1"], "outputs_a": ["Sheet1!A1"]},
            )

    def test_generate_entrypoints_emits_named_functions(self):
        """Generated code should include named entrypoints when configured."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
            _make_node("Sheet1!C1", "=Sheet1!A1*3", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(
            ["Sheet1!B1"],
            entrypoints={
                "outputs": ["Sheet1!B1", "Sheet1!C1"],
                "inputs-1": ["Sheet1!A1"],
            },
        )
        assert "TARGETS_OUTPUTS" in code
        assert "TARGETS_INPUTS_1" in code
        assert "def compute_outputs(inputs=None, *, ctx=None):" in code
        assert "def compute_inputs_1(inputs=None, *, ctx=None):" in code
        assert "def compute_all(inputs=None, *, ctx=None):" in code

    def test_generate_entrypoints_rejects_name_collisions(self):
        """Entrypoints with normalized name collisions should raise."""
        graph = _make_graph(
            _make_node("Sheet1!A1", None, 10.0),
            _make_node("Sheet1!B1", None, 20.0),
        )
        gen = CodeGenerator(graph)
        with pytest.raises(ValueError, match="normalize"):
            _ = gen.generate(
                ["Sheet1!A1"],
                entrypoints={
                    "outputs-a": ["Sheet1!A1"],
                    "outputs_a": ["Sheet1!B1"],
                },
            )

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
        assert compute_all({"Sheet1!A1": 7.0})["Sheet1!B1"] == 14.0

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
        """Changing an input should invalidate only dependent cached cells."""
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

        ctx.set_inputs({"Sheet1!A1": 7.0})
        assert "Sheet1!A1" not in ctx.cache
        assert "Sheet1!B1" not in ctx.cache
        assert "Sheet1!C1" not in ctx.cache
        assert "Sheet1!D1" in ctx.cache
        assert "Sheet1!E1" in ctx.cache

        result = compute_all(ctx=ctx)
        assert result["Sheet1!C1"] == 15.0
        assert result["Sheet1!E1"] == 4.0
        assert call_count == {"C1": 2, "E1": 1}

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

        assert result["Sheet1!B1:Sheet1!C1"].tolist() == [[20.0, 30.0]]
