"""Tests for the extractor module and DependencyGraph filter methods.

Tests cell key format normalization and discovery of formula cells.
Keys should use quoted sheet names when the sheet name contains spaces,
hyphens, or other special characters that require quoting in Excel formulas.

Examples:
    - Simple sheet: Sheet1!A1 (unquoted)
    - Sheet with space: 'Sales Data'!A1 (quoted)
    - Sheet with hyphen: 'Q4-Results'!A1 (quoted)
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import xlsxwriter

from excel_grapher import (
    create_dependency_graph,
    discover_formula_cells_in_rows,
    format_cell_key,
    needs_quoting,
)


class TestNeedsQuoting:
    """Tests for the needs_quoting helper function."""

    def test_simple_alphanumeric_no_quoting(self) -> None:
        """Simple alphanumeric sheet names don't need quoting."""
        assert needs_quoting("Sheet1") is False
        assert needs_quoting("Data") is False
        assert needs_quoting("MySheet123") is False

    def test_sheet_with_space_needs_quoting(self) -> None:
        """Sheet names with spaces need quoting."""
        assert needs_quoting("Sales Data") is True
        assert needs_quoting("Input Values") is True
        assert needs_quoting("Q4 Results") is True

    def test_sheet_with_hyphen_needs_quoting(self) -> None:
        """Sheet names with hyphens need quoting."""
        assert needs_quoting("Q4-Results") is True
        assert needs_quoting("2024-Budget") is True
        assert needs_quoting("Baseline - external") is True

    def test_sheet_with_underscore_no_quoting(self) -> None:
        """Sheet names with underscores don't need quoting."""
        assert needs_quoting("Sales_Data") is False
        assert needs_quoting("Q4_Results") is False


class TestFormatCellKey:
    """Tests for the format_cell_key helper function."""

    def test_simple_sheet_unquoted(self) -> None:
        """Simple sheet names should produce unquoted keys."""
        assert format_cell_key("Sheet1", "A", 1) == "Sheet1!A1"
        assert format_cell_key("Data", "B", 10) == "Data!B10"
        assert format_cell_key("MySheet", "AA", 100) == "MySheet!AA100"

    def test_sheet_with_space_quoted(self) -> None:
        """Sheet names with spaces should produce quoted keys."""
        assert format_cell_key("Sales Data", "A", 1) == "'Sales Data'!A1"
        assert format_cell_key("Input Values", "C", 5) == "'Input Values'!C5"

    def test_sheet_with_hyphen_quoted(self) -> None:
        """Sheet names with hyphens should produce quoted keys."""
        assert format_cell_key("Q4-Results", "A", 1) == "'Q4-Results'!A1"
        assert format_cell_key("2024-Budget", "B", 2) == "'2024-Budget'!B2"

    def test_complex_sheet_name_quoted(self) -> None:
        """Sheet names with multiple special characters should be quoted."""
        assert format_cell_key("Baseline - external", "M", 35) == "'Baseline - external'!M35"
        assert format_cell_key("GDP Forecast", "A", 1) == "'GDP Forecast'!A1"


class TestDiscoverFormulaCellsInRows:
    """Tests for discover_formula_cells_in_rows function."""

    def test_discovers_formula_cells_with_numeric_values(self, tmp_path: Path) -> None:
        """Should discover formula cells that have numeric cached values."""
        excel_path = tmp_path / "discover_test.xlsx"

        # Use xlsxwriter to create workbook with cached values
        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")

        # Row 1: value, formula with numeric result, formula with text result
        ws.write_number(0, 0, 10)  # A1 = 10 (value)
        ws.write_formula(0, 1, "=A1*2", None, 20)  # B1 = =A1*2 (cached: 20)
        ws.write_formula(0, 2, '="text"', None, "text")  # C1 = ="text" (cached: "text")

        # Row 2: more formulas with numeric results
        ws.write_number(1, 0, 5)  # A2 = 5 (value)
        ws.write_formula(1, 1, "=A2+10", None, 15)  # B2 = =A2+10 (cached: 15)

        wb.close()

        targets = discover_formula_cells_in_rows(excel_path, "Sheet1", [1, 2])

        # Should find B1 and B2 (formulas with numeric values)
        # Should NOT find A1, A2 (values) or C1 (formula with text result)
        assert "Sheet1!B1" in targets
        assert "Sheet1!B2" in targets
        assert "Sheet1!A1" not in targets
        assert "Sheet1!A2" not in targets
        assert "Sheet1!C1" not in targets

    def test_returns_empty_for_nonexistent_sheet(self, tmp_path: Path) -> None:
        """Should return empty list for non-existent sheet."""
        excel_path = tmp_path / "test.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "Sheet1"
        wb.save(excel_path)
        wb.close()

        targets = discover_formula_cells_in_rows(excel_path, "NonExistent", [1])
        assert targets == []

    def test_quoted_sheet_names(self, tmp_path: Path) -> None:
        """Should return properly quoted keys for sheets with special chars."""
        excel_path = tmp_path / "quoted_test.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sales Data")
        ws.write_number(0, 0, 100)  # A1 = 100
        ws.write_formula(0, 1, "=A1*2", None, 200)  # B1 = =A1*2
        wb.close()

        targets = discover_formula_cells_in_rows(excel_path, "Sales Data", [1])

        assert len(targets) == 1
        assert targets[0] == "'Sales Data'!B1"


class TestDependencyGraphFilterMethods:
    """Tests for DependencyGraph filter methods (formula_nodes, leaf_node_items, etc.)."""

    def test_formula_nodes_returns_formula_cells(self, tmp_path: Path) -> None:
        """formula_nodes() should iterate over nodes with formulas."""
        excel_path = tmp_path / "filter_test.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        ws.write_number(0, 0, 2)  # A1 = 2
        ws.write_number(1, 0, 3)  # A2 = 3
        ws.write_formula(2, 0, "=A1+A2", None, 5)  # A3 = =A1+A2
        wb.close()

        graph = create_dependency_graph(excel_path, ["Sheet1!A3"], load_values=True)
        formula_items = list(graph.formula_nodes())

        assert len(formula_items) == 1
        key, node = formula_items[0]
        assert key == "Sheet1!A3"
        assert node.formula == "=A1+A2"

    def test_leaf_node_items_returns_leaf_cells(self, tmp_path: Path) -> None:
        """leaf_node_items() should iterate over leaf nodes."""
        excel_path = tmp_path / "filter_test.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        ws.write_number(0, 0, 2)  # A1 = 2
        ws.write_number(1, 0, 3)  # A2 = 3
        ws.write_formula(2, 0, "=A1+A2", None, 5)  # A3 = =A1+A2
        wb.close()

        graph = create_dependency_graph(excel_path, ["Sheet1!A3"], load_values=True)
        leaf_items = list(graph.leaf_node_items())

        assert len(leaf_items) == 2
        keys = {key for key, _ in leaf_items}
        assert keys == {"Sheet1!A1", "Sheet1!A2"}

    def test_formula_keys_returns_sorted_list(self, tmp_path: Path) -> None:
        """formula_keys() should return sorted list of formula cell keys."""
        excel_path = tmp_path / "filter_test.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        ws.write_number(0, 0, 1)  # A1 = 1
        ws.write_formula(0, 1, "=A1*2", None, 2)  # B1 = =A1*2
        ws.write_formula(0, 2, "=B1+1", None, 3)  # C1 = =B1+1
        wb.close()

        graph = create_dependency_graph(excel_path, ["Sheet1!C1"], load_values=True)
        keys = graph.formula_keys()

        # Should be sorted alphabetically
        assert keys == ["Sheet1!B1", "Sheet1!C1"]

    def test_leaf_keys_returns_sorted_list(self, tmp_path: Path) -> None:
        """leaf_keys() should return sorted list of leaf node keys."""
        excel_path = tmp_path / "filter_test.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        ws.write_number(0, 0, 10)  # A1 = 10
        ws.write_number(0, 1, 20)  # B1 = 20
        ws.write_formula(0, 2, "=A1+B1", None, 30)  # C1 = =A1+B1
        wb.close()

        graph = create_dependency_graph(excel_path, ["Sheet1!C1"], load_values=True)
        keys = graph.leaf_keys()

        assert keys == ["Sheet1!A1", "Sheet1!B1"]

    def test_get_node_provides_direct_access(self, tmp_path: Path) -> None:
        """get_node() should provide O(1) access to node data."""
        excel_path = tmp_path / "access_test.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        ws.write_number(0, 0, 42)  # A1 = 42
        ws.write_formula(1, 0, "=A1*2", None, 84)  # A2 = =A1*2
        wb.close()

        graph = create_dependency_graph(excel_path, ["Sheet1!A2"], load_values=True)

        # Access formula node
        node = graph.get_node("Sheet1!A2")
        assert node is not None
        assert node.formula == "=A1*2"
        assert node.value == 84

        # Access value node
        node = graph.get_node("Sheet1!A1")
        assert node is not None
        assert node.formula is None
        assert node.value == 42

    def test_graph_preserves_structure(self, tmp_path: Path) -> None:
        """DependencyGraph should preserve edge information."""
        excel_path = tmp_path / "structure_test.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("Sheet1")
        ws.write_number(0, 0, 1)  # A1 = 1
        ws.write_number(0, 1, 2)  # B1 = 2
        ws.write_formula(0, 2, "=A1+B1", None, 3)  # C1 = =A1+B1
        wb.close()

        graph = create_dependency_graph(excel_path, ["Sheet1!C1"], load_values=True)

        # Check dependencies
        deps = graph.dependencies("Sheet1!C1")
        assert deps == {"Sheet1!A1", "Sheet1!B1"}

        # Check reverse edges (dependents)
        assert "Sheet1!C1" in graph.dependents("Sheet1!A1")
        assert "Sheet1!C1" in graph.dependents("Sheet1!B1")

    def test_handles_quoted_sheet_names(self, tmp_path: Path) -> None:
        """Should handle sheet names that require quoting."""
        excel_path = tmp_path / "quoted_sheet.xlsx"

        wb = xlsxwriter.Workbook(excel_path)
        ws = wb.add_worksheet("My Data")
        ws.write_number(0, 0, 5)  # A1 = 5
        ws.write_formula(1, 0, "=A1+10", None, 15)  # A2 = =A1+10
        wb.close()

        targets = discover_formula_cells_in_rows(excel_path, "My Data", [2])
        graph = create_dependency_graph(excel_path, targets, load_values=True)

        # Keys should be properly quoted
        assert graph.get_node("'My Data'!A2") is not None
        assert graph.get_node("'My Data'!A1") is not None

        # Verify the formula cell
        node = graph.get_node("'My Data'!A2")
        assert node.formula is not None
        assert node.value == 15

    def test_cross_sheet_references_correct_quoting(self, tmp_path: Path) -> None:
        """Should handle cross-sheet references with correct quoting."""
        excel_path = tmp_path / "cross_sheet.xlsx"

        wb = xlsxwriter.Workbook(excel_path)

        # Simple sheet (no quoting needed)
        ws1 = wb.add_worksheet("Data")
        ws1.write_number(0, 0, 100)  # Data!A1 = 100

        # Sheet with space (needs quoting)
        ws2 = wb.add_worksheet("Output Sheet")
        ws2.write_formula(0, 0, "=Data!A1*2", None, 200)  # 'Output Sheet'!A1 = =Data!A1*2

        wb.close()

        targets = discover_formula_cells_in_rows(excel_path, "Output Sheet", [1])
        graph = create_dependency_graph(excel_path, targets, load_values=True)

        # Data sheet key should NOT be quoted
        assert graph.get_node("Data!A1") is not None
        # Output Sheet key SHOULD be quoted
        assert graph.get_node("'Output Sheet'!A1") is not None

        # Verify proper quoting in all keys
        for key in graph:
            if "Output Sheet" in key:
                assert key.startswith("'Output Sheet'!"), f"Key {key} should be quoted"
            elif "Data" in key:
                assert not key.startswith("'Data'!"), f"Key {key} should NOT be quoted"
