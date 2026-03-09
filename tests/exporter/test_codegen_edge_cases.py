"""Tests for code generation edge cases."""

from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from excel_grapher import DependencyGraph, Node, create_dependency_graph
from excel_grapher.evaluator.codegen import CodeGenerator
from excel_grapher.evaluator.name_utils import parse_address


def _make_node(address: str, formula: str | None, value: object) -> Node:
    """Helper to create a Node from a sheet-qualified address.

    Note: Sheet names should NOT include quotes when stored on the Node,
    as Node.key will add them automatically when needed.
    """
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


class TestCircularReferences:
    """Tests for circular reference handling.

    Excel allows circular references when broken by conditional evaluation
    (IF, IFERROR, etc.) or with iterative calculation enabled. The code
    generator permits cycles in the dependency graph - actual circular
    reference errors will occur at runtime if the cycle is not broken.
    """

    def test_direct_circular_reference_generates_code(self):
        """Cell references itself - code generates but would error at runtime."""
        graph = _make_graph(
            _make_node("S!A1", "=S!A1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!A1"])

        # Code should generate successfully
        assert "def cell_s_a1(ctx):" in code
        assert "xl_eval(ctx, 'S!A1', cell_s_a1)" in code  # Self-reference in the formula

        # Code is valid Python (can be compiled)
        compile(code, "<string>", "exec")

    def test_indirect_circular_reference_generates_code(self):
        """A -> B -> A cycle generates valid code."""
        graph = _make_graph(
            _make_node("S!A1", "=S!B1", None),
            _make_node("S!B1", "=S!A1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!A1"])

        # Both cells should be in generated code
        assert "def cell_s_a1(ctx):" in code
        assert "def cell_s_b1(ctx):" in code

        # Code is valid Python
        compile(code, "<string>", "exec")

    def test_longer_circular_chain_generates_code(self):
        """A -> B -> C -> A cycle generates valid code."""
        graph = _make_graph(
            _make_node("S!A1", "=S!B1", None),
            _make_node("S!B1", "=S!C1", None),
            _make_node("S!C1", "=S!A1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!A1"])

        # All cells should be in generated code
        assert "def cell_s_a1(ctx):" in code
        assert "def cell_s_b1(ctx):" in code
        assert "def cell_s_c1(ctx):" in code

        # Code is valid Python
        compile(code, "<string>", "exec")

    def test_conditional_dependency_without_cycle(self):
        """IF with non-cyclic dependencies executes correctly."""
        graph = _make_graph(
            _make_node("S!A1", None, True),
            _make_node("S!B1", None, 10),
            _make_node("S!C1", None, 20),
            _make_node("S!D1", "=IF(S!A1,S!B1,S!C1)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!D1"])

        # Execute the code
        namespace: dict = {}
        exec(code, namespace)

        result = namespace["compute_all"]()
        assert result["S!D1"] == 10  # A1 is True, so B1 is returned


class TestMissingCellReferences:
    """Tests for formulas that reference cells not in the graph."""

    def test_missing_cell_reference_raises(self):
        """Formula references cell not in graph - evaluation raises."""
        graph = _make_graph(
            _make_node("S!A1", "=S!B1+1", None),
            # S!B1 is NOT in the graph
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!A1"])

        # Should evaluate missing refs via xl_cell(), not stub functions.
        assert "xl_cell(ctx, 'S!B1')" in code
        assert "def cell_s_b1" not in code

        # Should be executable (missing cell raises at runtime).
        namespace: dict = {}
        exec(code, namespace)
        with pytest.raises(KeyError):
            namespace["compute_all"]()

    def test_missing_cell_in_range_raises(self):
        """Range includes cells not in graph - evaluation raises."""
        graph = _make_graph(
            _make_node("S!A1", None, 1.0),
            # S!A2, S!A3 are NOT in the graph
            _make_node("S!B1", "=SUM(S!A1:S!A3)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!B1"])

        # Missing cells are evaluated via xl_cell() and must raise KeyError.
        assert "xl_cell(ctx, 'S!A2')" in code
        assert "xl_cell(ctx, 'S!A3')" in code

        namespace: dict = {}
        exec(code, namespace)
        with pytest.raises(KeyError):
            namespace["compute_all"]()

    def test_missing_cell_in_dynamic_offset_raises(self):
        """Dynamic OFFSET that lands on a missing cell should raise (do not conceal missing inputs)."""
        graph = _make_graph(
            _make_node("S!A1", None, 0),
            _make_node("S!C1", None, 1),  # dynamic row offset
            _make_node("S!B1", "=OFFSET(S!A1,S!C1,0)", None),  # -> S!A2
            # S!A2 is NOT in the graph
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!B1"])

        # Must use runtime OFFSET (row offset is not a literal constant).
        assert "xl_offset(ctx" in code

        namespace: dict = {}
        exec(code, namespace)
        with pytest.raises(KeyError):
            namespace["compute_all"]()


class TestSpecialCharactersInSheetNames:
    """Tests for sheet names with special characters."""

    def test_sheet_name_with_spaces(self):
        """Sheet name contains spaces."""
        graph = _make_graph(
            _make_node("'My Sheet'!A1", None, 100.0),
            _make_node("'My Sheet'!B1", "='My Sheet'!A1*2", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["'My Sheet'!B1"])

        assert "cell_my_sheet_b1" in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["'My Sheet'!B1"] == 200.0

    def test_sheet_name_with_numbers(self):
        """Sheet name is numeric."""
        graph = _make_graph(
            _make_node("'2024'!A1", None, 50.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["'2024'!A1"])

        assert "DEFAULT_INPUTS" in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        # Note: Address is normalized - quotes removed when not needed
        assert result["2024!A1"] == 50.0

    def test_sheet_name_with_special_chars(self):
        """Sheet name contains special characters."""
        graph = _make_graph(
            _make_node("'Data (v2)'!A1", None, 25.0),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["'Data (v2)'!A1"])

        assert "DEFAULT_INPUTS" in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        # Sheet name has space, so quotes are preserved
        assert result["'Data (v2)'!A1"] == 25.0


class TestComplexFormulas:
    """Tests for complex or unusual formulas."""

    def test_deeply_nested_functions(self):
        """Deeply nested function calls."""
        graph = _make_graph(
            _make_node("S!A1", None, 10.0),
            _make_node("S!B1", "=SUM(S!A1, SUM(S!A1, SUM(S!A1, S!A1)))", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!B1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        # 10 + (10 + (10 + 10)) = 10 + (10 + 20) = 10 + 30 = 40
        assert result["S!B1"] == 40.0

    def test_multiple_operators(self):
        """Formula with many operators."""
        graph = _make_graph(
            _make_node("S!A1", None, 2.0),
            _make_node("S!B1", "=S!A1+S!A1*S!A1-S!A1/S!A1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!B1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        # 2 + 2*2 - 2/2 = 2 + 4 - 1 = 5
        assert result["S!B1"] == 5.0

    def test_string_concatenation(self):
        """String concatenation with & operator."""
        graph = _make_graph(
            _make_node("S!A1", None, "Hello"),
            _make_node("S!B1", None, "World"),
            _make_node("S!C1", '=S!A1&" "&S!B1', None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!C1"] == "Hello World"

    def test_comparison_operators(self):
        """Comparison operators in formulas."""
        graph = _make_graph(
            _make_node("S!A1", None, 10.0),
            _make_node("S!B1", None, 5.0),
            _make_node("S!C1", "=IF(S!A1>S!B1, S!A1, S!B1)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!C1"] == 10.0


class TestErrorHandling:
    """Tests for error value handling."""

    def test_division_by_zero(self):
        """Division by zero returns XlError.DIV (Excel semantics)."""
        graph = _make_graph(
            _make_node("S!A1", None, 10.0),
            _make_node("S!B1", None, 0.0),
            _make_node("S!C1", "=S!A1/S!B1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])

        namespace: dict = {}
        exec(code, namespace)
        # xl_div returns XlError.DIV for division by zero (Excel semantics)
        result = namespace["compute_all"]()
        # Check that we got an XlError enum value (the generated code defines its own XlError)
        assert str(result["S!C1"]) == "XlError.DIV"

    def test_error_literal_in_formula(self):
        """Formula contains error literal."""
        graph = _make_graph(
            _make_node("S!A1", "=#N/A", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!A1"])

        assert "XlError.NA" in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!A1"] == namespace["XlError"].NA


class TestOffsetFunction:
    """Tests for OFFSET function handling."""

    def test_offset_static_single_cell(self):
        """OFFSET with constant offsets resolves statically to single cell."""
        graph = _make_graph(
            _make_node("S!A1", None, 10.0),
            _make_node("S!B2", None, 20.0),
            _make_node("S!C3", None, 30.0),
            # OFFSET(A1, 1, 1) -> B2
            _make_node("S!D1", "=OFFSET(S!A1,1,1)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!D1"])

        # Should NOT have _CELL_TABLE definition (static resolution)
        assert "_CELL_TABLE = {" not in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!D1"] == 20.0

    def test_offset_static_range(self):
        """OFFSET with constant offsets and size resolves statically to range."""
        graph = _make_graph(
            _make_node("S!A1", None, 1.0),
            _make_node("S!B1", None, 2.0),
            _make_node("S!A2", None, 3.0),
            _make_node("S!B2", None, 4.0),
            # OFFSET(A1, 0, 0, 2, 2) -> A1:B2 as 2D array
            _make_node("S!C1", "=SUM(OFFSET(S!A1,0,0,2,2))", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])

        # Should NOT have _CELL_TABLE definition (static resolution)
        assert "_CELL_TABLE = {" not in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!C1"] == 10.0  # 1+2+3+4

    def test_offset_dynamic_single_cell(self):
        """OFFSET with dynamic offset uses runtime lookup."""
        graph = _make_graph(
            _make_node("S!A1", None, 10.0),
            _make_node("S!A2", None, 20.0),
            _make_node("S!A3", None, 30.0),
            _make_node("S!B1", None, 1.0),  # Dynamic offset value
            # OFFSET(A1, B1, 0) -> A2 (when B1=1)
            _make_node("S!C1", "=OFFSET(S!A1,S!B1,0)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])

        # Dynamic OFFSET should not emit a cell table.
        assert "_CELL_TABLE = {" not in code
        assert "xl_offset" in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!C1"] == 20.0  # A2

    def test_offset_dynamic_range(self):
        """OFFSET with dynamic size uses runtime lookup."""
        graph = _make_graph(
            _make_node("S!A1", None, 1.0),
            _make_node("S!A2", None, 2.0),
            _make_node("S!A3", None, 3.0),
            _make_node("S!B1", None, 2.0),  # Dynamic height
            # OFFSET(A1, 0, 0, B1, 1) -> A1:A2 (when B1=2)
            _make_node("S!C1", "=SUM(OFFSET(S!A1,0,0,S!B1,1))", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])

        # Dynamic OFFSET should not emit a cell table.
        assert "_CELL_TABLE = {" not in code

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!C1"] == 3.0  # 1+2

    def test_offset_negative_offset(self):
        """OFFSET with negative offset goes backwards."""
        graph = _make_graph(
            _make_node("S!A1", None, 10.0),
            _make_node("S!A2", None, 20.0),
            _make_node("S!A3", None, 30.0),
            # OFFSET(A3, -2, 0) -> A1
            _make_node("S!B1", "=OFFSET(S!A3,-2,0)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!B1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!B1"] == 10.0

    def test_offset_invalid_reference(self):
        """OFFSET that results in invalid reference returns error."""
        graph = _make_graph(
            _make_node("S!A1", None, 10.0),
            # OFFSET(A1, -1, 0) -> invalid (row 0)
            _make_node("S!B1", "=OFFSET(S!A1,-1,0)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!B1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert str(result["S!B1"]) == "XlError.REF"

    def test_offset_with_formulas(self):
        """OFFSET works with formula cells."""
        graph = _make_graph(
            _make_node("S!A1", None, 5.0),
            _make_node("S!A2", "=S!A1*2", None),  # 10
            _make_node("S!A3", "=S!A2*2", None),  # 20
            # OFFSET(A1, 2, 0) -> A3
            _make_node("S!B1", "=OFFSET(S!A1,2,0)", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!B1"])

        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!B1"] == 20.0

    def test_offset_dynamic_base_range(self):
        """Dynamic OFFSET supports range base references and inherits base size when height/width are omitted."""
        graph = _make_graph(
            _make_node("S!A1", None, 1.0),
            _make_node("S!A2", None, 2.0),
            _make_node("S!A3", None, 3.0),
            _make_node("S!B1", None, 1.0),
            # Excel semantics: base is the top-left of the range, and omitted height/width
            # inherit the base range size. This should sum OFFSET(A1:A2, 1, 0) = A2:A3.
            _make_node("S!C1", "=SUM(OFFSET(S!A1:S!A2,S!B1,0))", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])
        namespace: dict = {}
        exec(code, namespace)
        result = namespace["compute_all"]()
        assert result["S!C1"] == 5.0


class TestNamedRangeRegression:
    """Regression tests for named range handling in code generation."""

    def test_vlookup_with_range_named_range_is_exportable(
        self,
        tmp_path: Path,
    ) -> None:
        """VLOOKUP over a range-based named range should not crash codegen."""
        excel_path = tmp_path / "named_range_vlookup.xlsx"

        wb = openpyxl.Workbook()
        ws_table = wb.active
        ws_table.title = "Table"
        ws_table["A1"].value = 1
        ws_table["B1"].value = 10

        ws_chart = wb.create_sheet("Chart Data")
        ws_chart["D11"].value = 1
        ws_chart["E11"].value = (
            "=VLOOKUP('Chart Data'!D11, NumRiskTable, 2, FALSE())"
        )

        wb.defined_names.add(
            DefinedName("NumRiskTable", attr_text="Table!$A$1:$B$1")
        )
        wb.save(excel_path)

        graph = create_dependency_graph(
            excel_path,
            ["'Chart Data'!E11"],
            load_values=False,
        )
        gen = CodeGenerator(graph)

        # GREEN: After normalization, NumRiskTable is expanded to a sheet-qualified
        # range in normalized_formula, so core.formula_ast.parse succeeds and
        # CodeGenerator can emit a runnable package without raising ParseError.
        files = gen.generate_modules(["'Chart Data'!E11"])
        assert "exported/entrypoint.py" in files

    def test_offset_dynamic_cell_table_excludes_unreachable_sheets(self):
        """Dynamic OFFSET should not force _CELL_TABLE to include unrelated sheets."""
        graph = _make_graph(
            # Sheet S: used by OFFSET base reference
            _make_node("S!A1", None, 10.0),
            _make_node("S!A2", None, 20.0),
            _make_node("S!B1", None, 1.0),
            _make_node("S!C1", "=OFFSET(S!A1,S!B1,0)", None),
            # Sheet T: completely unrelated to any OFFSET base reference
            _make_node("T!A1", None, 999.0),
            _make_node("T!B1", "=T!A1+1", None),
        )
        gen = CodeGenerator(graph)
        code = gen.generate(["S!C1"])

        assert "_CELL_TABLE = {" not in code
