"""Tests for name_utils module."""

import pytest

from excel_grapher.core.address_keys import (
    format_key as format_address,
)
from excel_grapher.core.address_keys import (
    normalize_key as normalize_address,
)
from excel_grapher.core.address_keys import (
    parse_address,
    quote_sheet_if_needed,
)
from excel_grapher.core.excel_function_names import (
    excel_func_to_python_runtime_name,
    excel_function_call_prefixes,
)
from excel_grapher.evaluator.name_utils import normalize_excel_function_name


def _excel_func_to_python(name: str) -> str:
    return excel_func_to_python_runtime_name(normalize_excel_function_name(name))


class TestNormalizeExcelFunctionName:
    """Tests for normalize_excel_function_name."""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("IFNA", "IFNA"),
            ("ifna", "IFNA"),
            ("_XLFN.IFNA", "IFNA"),
            ("_xlfn.IFNA", "IFNA"),
            ("_XLUDF.IFNA", "IFNA"),
            ("_xludf.XLOOKUP", "XLOOKUP"),
            ("_XLFN.NUMBERVALUE", "NUMBERVALUE"),
            ("_XLUDF.NUMBERVALUE", "NUMBERVALUE"),
            ("SUM", "SUM"),
            ("_XLFN.IFS", "IFS"),
            ("NORM.DIST", "NORM.DIST"),
            ("_XLUDF.MYADDIN", "_XLUDF.MYADDIN"),
            ("_XLUDF.CUSTOM_UDF", "_XLUDF.CUSTOM_UDF"),
        ],
    )
    def test_normalize_excel_function_name(self, raw: str, expected: str) -> None:
        assert normalize_excel_function_name(raw) == expected

    def test_excel_func_to_python_preserves_unknown_xludf_addin_name(self) -> None:
        assert _excel_func_to_python("_XLUDF.MYADDIN") == "xl__xludf_myaddin"

    def test_excel_function_call_prefixes_includes_compatibility_variants(self) -> None:
        assert excel_function_call_prefixes("IFNA") == (
            "IFNA(",
            "_XLFN.IFNA(",
            "_XLUDF.IFNA(",
        )

    def test_excel_function_call_prefixes_includes_xludf_for_all_functions(self) -> None:
        assert excel_function_call_prefixes("SUM") == (
            "SUM(",
            "_XLFN.SUM(",
            "_XLUDF.SUM(",
        )


class TestExcelFuncToPython:
    """Tests for Excel function names mapped through the export-runtime helper."""

    def test_simple_function(self):
        """Simple function name."""
        assert _excel_func_to_python("SUM") == "xl_sum"

    def test_multi_word_function(self):
        """Multi-word function name."""
        assert _excel_func_to_python("VLOOKUP") == "xl_vlookup"

    def test_function_with_numbers(self):
        """Function name with numbers."""
        assert _excel_func_to_python("LOG10") == "xl_log10"

    def test_already_lowercase(self):
        """Function name that's already lowercase (edge case)."""
        assert _excel_func_to_python("sum") == "xl_sum"

    def test_mixed_case(self):
        """Mixed case function name."""
        assert _excel_func_to_python("SumProduct") == "xl_sumproduct"

    def test_function_with_dot(self):
        """Function with dot (e.g., NORM.DIST)."""
        assert _excel_func_to_python("NORM.DIST") == "xl_norm_dist"

    def test_function_with_underscore(self):
        """Function with underscore (rare but possible)."""
        assert _excel_func_to_python("AGGREGATE_X") == "xl_aggregate_x"

    def test_prefixed_function_normalizes_before_python_name(self):
        """Compatibility prefixes should not appear in generated Python names."""
        assert _excel_func_to_python("_XLUDF.IFNA") == "xl_ifna"
        assert _excel_func_to_python("_xlfn.XLOOKUP") == "xl_xlookup"
        assert _excel_func_to_python("_xlfn.NUMBERVALUE") == "xl_numbervalue"
        assert _excel_func_to_python("SUM") == _excel_func_to_python("_xlfn.SUM")


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
        assert quote_sheet_if_needed("It's Data") == "'It''s Data'"

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
        assert format_address("It's Data", "C3") == "'It''s Data'!C3"


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
        assert normalize_address("'It''s Data'!C3") == "'It''s Data'!C3"
