"""Integration tests for CodeGenerator - verifies generated code matches evaluator."""

from __future__ import annotations

import time
from pathlib import Path

import pytest
from excel_grapher import DependencyGraph, create_dependency_graph
from excel_grapher import CycleError, Node

from excel_grapher.evaluator.codegen import CodeGenerator
from excel_grapher.evaluator.name_utils import normalize_address
from excel_grapher.evaluator.name_utils import parse_address
from tests.evaluator.discover_formula_cells import discover_formula_cells_in_rows
from tests.evaluator.parity_harness import assert_codegen_matches_evaluator, exec_generated_code


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


class TestGeneratedCodeExecution:
    """Tests that generated code executes and produces correct results."""

    def test_generated_code_percent_literal(self) -> None:
        """Excel percent literals like 100% should parse and evaluate as 1.0."""
        graph = _make_graph(
            _make_node("S!A1", "=100%/29", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1"], rtol=1e-12, atol=0.0)
        assert result.generated_results["S!A1"] == pytest.approx(1.0 / 29.0, rel=1e-12)

    def test_generated_code_xlookup_matches_evaluator(self) -> None:
        """Generated code matches evaluator for XLOOKUP (including _xlfn prefix)."""
        graph = _make_graph(
            _make_node("S!A1", None, 1),
            _make_node("S!A2", None, 2),
            _make_node("S!A3", None, 3),
            _make_node("S!B1", None, "a"),
            _make_node("S!B2", None, "b"),
            _make_node("S!B3", None, "c"),
            _make_node("S!C1", "=_xlfn.XLOOKUP(2,S!A1:S!A3,S!B1:S!B3)", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!C1"])
        assert result.generated_results["S!C1"] == "b"

    def test_generated_code_simple_arithmetic(self) -> None:
        """Generated code correctly computes simple arithmetic."""
        graph = _make_graph(
            _make_node("S!A1", None, 10),
            _make_node("S!A2", None, 5),
            _make_node("S!B1", "=S!A1+S!A2", None),
        )

        generated_results, _code, _ns = exec_generated_code(graph, ["S!B1"])
        assert generated_results["S!B1"] == 15.0

    def test_generated_code_matches_evaluator(self) -> None:
        """Generated code produces same results as FormulaEvaluator."""
        graph = _make_graph(
            _make_node("S!A1", None, 100),
            _make_node("S!A2", None, 50),
            _make_node("S!B1", "=S!A1*2", None),
            _make_node("S!B2", "=S!A2*3", None),
            _make_node("S!C1", "=S!B1+S!B2", None),
        )

        targets = ["S!C1", "S!B1", "S!B2"]
        assert_codegen_matches_evaluator(graph, targets)

    def test_generated_code_with_functions(self) -> None:
        """Generated code correctly handles Excel functions."""
        graph = _make_graph(
            _make_node("S!A1", None, 10),
            _make_node("S!A2", None, 20),
            _make_node("S!A3", None, 30),
            _make_node("S!B1", "=SUM(S!A1,S!A2,S!A3)", None),
            _make_node("S!B2", "=AVERAGE(S!A1,S!A2,S!A3)", None),
            _make_node("S!B3", "=MAX(S!A1,S!A2,S!A3)", None),
            _make_node("S!B4", "=MIN(S!A1,S!A2,S!A3)", None),
        )

        targets = ["S!B1", "S!B2", "S!B3", "S!B4"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.generated_results["S!B1"] == 60.0
        assert result.generated_results["S!B2"] == 20.0
        assert result.generated_results["S!B3"] == 30.0
        assert result.generated_results["S!B4"] == 10.0

    def test_generated_code_with_if(self) -> None:
        """Generated code correctly handles IF function."""
        graph = _make_graph(
            _make_node("S!A1", None, 10),
            _make_node("S!A2", None, 5),
            _make_node("S!B1", "=IF(S!A1>S!A2,S!A1,S!A2)", None),
            _make_node("S!B2", "=IF(S!A1<S!A2,S!A1,S!A2)", None),
        )

        targets = ["S!B1", "S!B2"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.generated_results["S!B1"] == 10
        assert result.generated_results["S!B2"] == 5

    def test_generated_code_if_uses_excel_boolean_coercion(self) -> None:
        """IF should use Excel-style boolean coercion (e.g., 'FALSE' -> False)."""
        graph = _make_graph(
            _make_node("S!A1", None, "FALSE"),
            _make_node("S!B1", "=IF(S!A1,1,2)", None),
        )

        targets = ["S!B1"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.evaluator_results["S!B1"] == 2
        assert result.generated_results["S!B1"] == 2

    def test_generated_code_comparison_matches_evaluator_for_numeric_strings(self) -> None:
        """Comparisons should coerce numeric strings like the evaluator (e.g., '0' = 0)."""
        graph = _make_graph(
            _make_node("S!A1", None, "0"),
            _make_node("S!A2", None, 0),
            _make_node("S!B1", "=S!A1=S!A2", None),
            _make_node("S!B2", "=S!A1<S!A2", None),
        )

        targets = ["S!B1", "S!B2"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.generated_results == result.evaluator_results

    def test_generated_code_with_nested_functions(self) -> None:
        """Generated code handles nested function calls."""
        graph = _make_graph(
            _make_node("S!A1", None, 5),
            _make_node("S!A2", None, 10),
            _make_node("S!A3", None, 15),
            _make_node("S!B1", "=MAX(S!A1,S!A2,S!A3)", None),
            _make_node("S!B2", "=ROUND(S!B1/S!A1,2)", None),  # round(15/5, 2) = 3.0
        )

        targets = ["S!B2"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.generated_results["S!B2"] == 3.0

    def test_generated_code_with_range(self) -> None:
        """Generated code correctly expands ranges."""
        graph = _make_graph(
            _make_node("S!A1", None, 1),
            _make_node("S!A2", None, 2),
            _make_node("S!A3", None, 3),
            _make_node("S!B1", "=SUM(S!A1:S!A3)", None),
        )

        targets = ["S!B1"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.generated_results["S!B1"] == 6.0

    def test_generated_code_with_string_operations(self) -> None:
        """Generated code handles string concatenation."""
        graph = _make_graph(
            _make_node("S!A1", None, "Hello"),
            _make_node("S!A2", None, "World"),
            _make_node("S!B1", '=S!A1&" "&S!A2', None),
        )

        targets = ["S!B1"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.generated_results["S!B1"] == "Hello World"

    def test_generated_code_with_comparison(self) -> None:
        """Generated code handles comparison operators."""
        graph = _make_graph(
            _make_node("S!A1", None, 10),
            _make_node("S!A2", None, 10),
            _make_node("S!B1", "=S!A1=S!A2", None),
            _make_node("S!B2", "=S!A1<>S!A2", None),
        )

        targets = ["S!B1", "S!B2"]
        result = assert_codegen_matches_evaluator(graph, targets)
        assert result.generated_results["S!B1"] is True
        assert result.generated_results["S!B2"] is False


class TestGeneratedCodeWithRealWorkbook:
    """Integration tests using real LIC-DSF workbook."""

    WORKBOOK_PATH = Path("example/data/lic-dsf-template-2025-08-12.xlsm")
    INDICATOR_CONFIG = {"B1_GDP_ext": [35, 36]}
    MAX_DEPTH = 100
    RTOL = 1e-5

    @pytest.mark.slow
    def test_generated_code_matches_evaluator_for_lic_dsf(self) -> None:
        """Generated code produces same results as evaluator for real workbook."""
        if not self.WORKBOOK_PATH.exists():
            pytest.skip(f"Test workbook not found at {self.WORKBOOK_PATH}")

        print("\n\nLoading graph...")
        start = time.time()
        targets: list[str] = []
        for sheet_name, rows in self.INDICATOR_CONFIG.items():
            targets.extend(
                discover_formula_cells_in_rows(self.WORKBOOK_PATH, sheet_name, rows)
            )
        print(f"Discovered {len(targets)} targets in {time.time() - start:.1f}s")

        start = time.time()
        graph = create_dependency_graph(
            self.WORKBOOK_PATH,
            targets,
            load_values=True,
            max_depth=self.MAX_DEPTH,
        )
        print(f"Built graph with {len(graph)} nodes in {time.time() - start:.1f}s")

        assert_codegen_matches_evaluator(
            graph,
            targets,
            rtol=self.RTOL,
            dependency_order=True,
            fail_fast=True,
        )

    @pytest.mark.slow
    def test_generated_code_is_standalone(self) -> None:
        """Generated code runs without excel_evaluator imports at runtime."""
        if not self.WORKBOOK_PATH.exists():
            pytest.skip(f"Test workbook not found at {self.WORKBOOK_PATH}")

        targets: list[str] = []
        for sheet_name, rows in self.INDICATOR_CONFIG.items():
            targets.extend(
                discover_formula_cells_in_rows(self.WORKBOOK_PATH, sheet_name, rows)
            )

        graph = create_dependency_graph(
            self.WORKBOOK_PATH,
            targets[:5],  # Just use first 5 targets for speed
            load_values=True,
            max_depth=self.MAX_DEPTH,
        )

        gen = CodeGenerator(graph)
        code = gen.generate(targets[:5])

        # Verify no excel_evaluator imports in generated code
        assert "from excel_evaluator" not in code
        assert "import excel_evaluator" not in code

        # Execute in clean namespace (no excel_evaluator available)
        namespace: dict = {"__builtins__": __builtins__}
        exec(code, namespace)

        # Should run without ImportError
        result = namespace["compute_all"]()
        assert isinstance(result, dict)
        assert len(result) > 0

    @pytest.mark.slow
    def test_generated_code_can_evaluate_dependency_closure_eagerly(self) -> None:
        """Generated code can evaluate the full dependency closure without missing-cell KeyErrors.

        This is an "eager sweep" of the dependency closure for a fixed cached-input assignment:
        we force evaluation of every formula cell reachable from the targets.
        """
        if not self.WORKBOOK_PATH.exists():
            pytest.skip(f"Test workbook not found at {self.WORKBOOK_PATH}")

        targets: list[str] = []
        for sheet_name, rows in self.INDICATOR_CONFIG.items():
            targets.extend(
                discover_formula_cells_in_rows(self.WORKBOOK_PATH, sheet_name, rows)
            )

        graph = create_dependency_graph(
            self.WORKBOOK_PATH,
            targets,
            load_values=True,
            max_depth=self.MAX_DEPTH,
        )

        # Compute dependency closure (graph-native). Normalize addresses to match generated runtime.
        closure: set[str] = set()
        stack = [normalize_address(t) for t in targets]
        while stack:
            addr = stack.pop()
            if addr in closure:
                continue
            node = graph.get_node(addr)
            if node is None:
                continue
            closure.add(addr)
            for dep in graph.dependencies(addr):
                dep_n = normalize_address(dep)
                if dep_n not in closure and graph.get_node(dep_n) is not None:
                    stack.append(dep_n)

        # Build a cached leaf-input mapping for the closure.
        inputs: dict[str, object] = {}
        for leaf in graph.leaf_keys():
            key = normalize_address(leaf)
            if key not in closure:
                continue
            node = graph.get_node(key)
            if node is None:
                continue
            inputs[key] = 0 if node.value is None else node.value

        # Generate the export and force evaluation of all formula cells in the closure.
        code = CodeGenerator(graph).generate(targets)
        namespace: dict = {"__builtins__": __builtins__}
        exec(code, namespace)

        make_context = namespace["make_context"]
        xl_cell = namespace["xl_cell"]
        ctx = make_context(inputs)

        # Prefer the graph's evaluation order (non-strict excludes may-cycle nodes).
        try:
            ordered = graph.evaluation_order(strict=False)
        except CycleError:
            ordered = []

        formula_targets: list[str] = []
        seen: set[str] = set()
        for addr in ordered:
            a = normalize_address(addr)
            if a in seen or a not in closure:
                continue
            seen.add(a)
            node = graph.get_node(a)
            if node is not None and node.formula is not None:
                formula_targets.append(a)

        # Include any remaining formula cells not covered by evaluation_order (e.g. excluded by cycles).
        for addr in sorted(closure):
            if addr in seen:
                continue
            node = graph.get_node(addr)
            if node is not None and node.formula is not None:
                formula_targets.append(addr)

        for addr in formula_targets:
            xl_cell(ctx, addr)
