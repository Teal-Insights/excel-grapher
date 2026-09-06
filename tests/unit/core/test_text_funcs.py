"""Tests for shared text worksheet function semantics."""

from excel_grapher.core import XlError
from excel_grapher.core.text_funcs import text_format


def test_text_format_propagates_format_argument_error() -> None:
    """TEXT returns an error passed as its format argument."""
    assert text_format(123.0, XlError.REF) is XlError.REF
