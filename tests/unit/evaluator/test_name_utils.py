"""Tests for name_utils module."""

import pytest

from excel_grapher.evaluator.name_utils import (
    address_to_python_name,
    excel_func_to_python,
    format_address,
    normalize_address,
    parse_address,
    quote_sheet_if_needed,
)


class TestAddressToPythonName:
    """Tests for address_to_python_name function."""

    def test_simple_address(self):
        """Simple sheet and cell reference."""
        assert address_to_python_name("Sheet1!A1") == "cell_sheet1_a1"

    def test_lowercase_preserved(self):
        """Result should be lowercase."""
        assert address_to_python_name("SHEET1!A1") == "cell_sheet1_a1"

    def test_quoted_sheet_name(self):
        """Sheet name with quotes (contains space)."""
        assert address_to_python_name("'My Sheet'!B2") == "cell_my_sheet_b2"

    def test_quoted_sheet_name_with_escaped_quote(self):
        """Sheet name with escaped single quote."""
        assert address_to_python_name("'It''s Data'!A1") == "cell_its_data_a1"

    def test_sheet_name_with_underscore(self):
        """Sheet name already containing underscores."""
        assert address_to_python_name("B1_GDP_ext!A35") == "cell_b1_gdp_ext_a35"

    def test_sheet_name_with_spaces(self):
        """Sheet name with multiple spaces."""
        assert address_to_python_name("'Sheet With Spaces'!C10") == "cell_sheet_with_spaces_c10"

    def test_multi_letter_column(self):
        """Column with multiple letters."""
        assert address_to_python_name("Sheet1!AA100") == "cell_sheet1_aa100"

    def test_special_characters_removed(self):
        """Special characters should become underscores or be removed."""
        assert address_to_python_name("'Data-2024'!A1") == "cell_data_2024_a1"

    def test_multiple_underscores_collapsed(self):
        """Multiple consecutive underscores should be collapsed."""
        assert address_to_python_name("'My  Sheet'!A1") == "cell_my_sheet_a1"

    def test_leading_underscore_after_cell_prefix(self):
        """No leading underscore after 'cell_' prefix."""
        # Sheet name starting with special char
        assert address_to_python_name("'_Hidden'!A1") == "cell_hidden_a1"

    def test_numeric_sheet_name(self):
        """Sheet name that is numeric (needs quotes in Excel)."""
        assert address_to_python_name("'2024'!A1") == "cell_2024_a1"

    def test_parentheses_in_sheet_name(self):
        """Sheet name with parentheses."""
        assert address_to_python_name("'Data (v2)'!B5") == "cell_data_v2_b5"


class TestExcelFuncToPython:
    """Tests for excel_func_to_python function."""

    def test_simple_function(self):
        """Simple function name."""
        assert excel_func_to_python("SUM") == "xl_sum"

    def test_multi_word_function(self):
        """Multi-word function name."""
        assert excel_func_to_python("VLOOKUP") == "xl_vlookup"

    def test_function_with_numbers(self):
        """Function name with numbers."""
        assert excel_func_to_python("LOG10") == "xl_log10"

    def test_already_lowercase(self):
        """Function name that's already lowercase (edge case)."""
        assert excel_func_to_python("sum") == "xl_sum"

    def test_mixed_case(self):
        """Mixed case function name."""
        assert excel_func_to_python("SumProduct") == "xl_sumproduct"

    def test_function_with_dot(self):
        """Function with dot (e.g., NORM.DIST)."""
        assert excel_func_to_python("NORM.DIST") == "xl_norm_dist"

    def test_function_with_underscore(self):
        """Function with underscore (rare but possible)."""
        assert excel_func_to_python("AGGREGATE_X") == "xl_aggregate_x"


class TestParseAddress:
    """Tests for parse_address function."""

    def test_simple_address(self):
        """Parse simple unquoted address."""
        sheet, cell = parse_address("Sheet1!A1")
        assert sheet == "Sheet1"
        assert cell == "A1"

    def test_quoted_sheet_with_space(self):
        """Parse address with quoted sheet containing space."""
        sheet, cell = parse_address("'My Sheet'!B2")
        assert sheet == "My Sheet"
        assert cell == "B2"

    def test_quoted_sheet_with_escaped_quote(self):
        """Parse address with escaped single quote in sheet name."""
        sheet, cell = parse_address("'It''s Data'!C3")
        assert sheet == "It's Data"
        assert cell == "C3"

    def test_quoted_numeric_sheet(self):
        """Parse address with quoted numeric sheet name."""
        sheet, cell = parse_address("'2024'!A1")
        assert sheet == "2024"
        assert cell == "A1"

    def test_invalid_no_exclamation(self):
        """Raise error for address without exclamation mark."""
        with pytest.raises(ValueError, match="sheet-qualified"):
            parse_address("NoExclamation")

    def test_invalid_quoted_no_exclamation(self):
        """Raise error for quoted address without exclamation mark."""
        with pytest.raises(ValueError, match="Invalid address"):
            parse_address("'Sheet Name'NoExclamation")


class TestQuoteSheetIfNeeded:
    """Tests for quote_sheet_if_needed function."""

    def test_simple_sheet_no_quotes(self):
        """Simple sheet name doesn't need quotes."""
        assert quote_sheet_if_needed("Sheet1") == "Sheet1"

    def test_sheet_with_space_needs_quotes(self):
        """Sheet with space needs quotes."""
        assert quote_sheet_if_needed("My Sheet") == "'My Sheet'"

    def test_sheet_with_hyphen_needs_quotes(self):
        """Sheet with hyphen needs quotes."""
        assert quote_sheet_if_needed("Data-2024") == "'Data-2024'"

    def test_sheet_with_apostrophe_needs_quotes(self):
        """Sheet with apostrophe needs quotes."""
        assert quote_sheet_if_needed("It's Data") == "'It's Data'"

    def test_numeric_sheet_no_quotes(self):
        """Purely numeric sheet doesn't need quotes."""
        assert quote_sheet_if_needed("2024") == "2024"

    def test_sheet_with_parentheses_no_quotes(self):
        """Sheet with parentheses (but no space/hyphen/apostrophe) doesn't need quotes."""
        assert quote_sheet_if_needed("Data(v2)") == "Data(v2)"


class TestFormatAddress:
    """Tests for format_address function."""

    def test_simple_address(self):
        """Format simple address without quotes."""
        assert format_address("Sheet1", "A1") == "Sheet1!A1"

    def test_address_needing_quotes(self):
        """Format address where sheet needs quotes."""
        assert format_address("My Sheet", "B2") == "'My Sheet'!B2"

    def test_address_with_apostrophe(self):
        """Format address with apostrophe in sheet name."""
        assert format_address("It's Data", "C3") == "'It's Data'!C3"


class TestNormalizeAddress:
    """Tests for normalize_address function."""

    def test_already_normalized(self):
        """Address already in normalized form stays the same."""
        assert normalize_address("Sheet1!A1") == "Sheet1!A1"

    def test_removes_unnecessary_quotes(self):
        """Remove quotes when not needed."""
        assert normalize_address("'2024'!A1") == "2024!A1"
        assert normalize_address("'Sheet1'!A1") == "Sheet1!A1"

    def test_keeps_needed_quotes(self):
        """Keep quotes when needed."""
        assert normalize_address("'My Sheet'!A1") == "'My Sheet'!A1"
        assert normalize_address("'Data-2024'!B2") == "'Data-2024'!B2"

    def test_handles_escaped_quotes(self):
        """Properly handle escaped quotes in sheet names."""
        assert normalize_address("'It''s Data'!C3") == "'It's Data'!C3"
