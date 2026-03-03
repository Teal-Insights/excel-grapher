"""Tests for parser.py split_top_level_* functions."""
from __future__ import annotations

from excel_grapher.grapher.parser import (
    split_top_level_choose,
    split_top_level_if,
    split_top_level_ifs,
    split_top_level_switch,
)


class TestSplitTopLevelIf:
    """Tests for split_top_level_if function."""

    def test_simple_if(self) -> None:
        """Basic IF formula should parse correctly."""
        result = split_top_level_if("=IF(A1>0, B1, C1)")
        assert result == ("A1>0", "B1", "C1")

    def test_if_with_trailing_addition_returns_none(self) -> None:
        """IF with trailing +D1 is NOT a top-level IF and should return None."""
        result = split_top_level_if("=IF(A1>0, B1, C1)+D1")
        assert result is None

    def test_if_with_trailing_multiplication_returns_none(self) -> None:
        """IF with trailing *D1 is NOT a top-level IF."""
        result = split_top_level_if("=IF(A1>0, B1, C1)*D1")
        assert result is None

    def test_if_with_trailing_cell_ref_returns_none(self) -> None:
        """Complex formula from bug report should return None."""
        formula = "=IF(ISNUMBER('Input 3 - Macro-Debt data(DMX)'!W51), '...'!W51, 0)+'Input 3 - Macro-Debt data(DMX)'!W205"
        result = split_top_level_if(formula)
        assert result is None

    def test_if_with_nested_parens(self) -> None:
        """IF with nested function calls should parse correctly."""
        result = split_top_level_if("=IF(SUM(A1:A10)>0, MAX(B1:B10), 0)")
        assert result == ("SUM(A1:A10)>0", "MAX(B1:B10)", "0")

    def test_if_with_trailing_whitespace(self) -> None:
        """IF with trailing whitespace only should still parse."""
        result = split_top_level_if("=IF(A1>0, B1, C1)   ")
        assert result == ("A1>0", "B1", "C1")

    def test_if_with_leading_whitespace(self) -> None:
        """IF with leading whitespace should parse."""
        result = split_top_level_if("=  IF(A1>0, B1, C1)")
        assert result == ("A1>0", "B1", "C1")

    def test_not_if_formula(self) -> None:
        """Non-IF formula should return None."""
        result = split_top_level_if("=SUM(A1:A10)")
        assert result is None


class TestSplitTopLevelIfs:
    """Tests for split_top_level_ifs function."""

    def test_simple_ifs(self) -> None:
        """Basic IFS formula should parse correctly."""
        result = split_top_level_ifs("=IFS(A1>0, B1, A1<0, C1)")
        assert result == ["A1>0", "B1", "A1<0", "C1"]

    def test_ifs_with_trailing_addition_returns_none(self) -> None:
        """IFS with trailing +D1 is NOT a top-level IFS."""
        result = split_top_level_ifs("=IFS(A1>0, B1, A1<0, C1)+D1")
        assert result is None

    def test_ifs_with_xlfn_prefix(self) -> None:
        """IFS with _xlfn. prefix should parse correctly."""
        result = split_top_level_ifs("=_xlfn.IFS(A1>0, B1)")
        assert result == ["A1>0", "B1"]

    def test_ifs_with_xlfn_and_trailing_returns_none(self) -> None:
        """IFS with _xlfn. prefix and trailing content should return None."""
        result = split_top_level_ifs("=_xlfn.IFS(A1>0, B1)+C1")
        assert result is None


class TestSplitTopLevelChoose:
    """Tests for split_top_level_choose function."""

    def test_simple_choose(self) -> None:
        """Basic CHOOSE formula should parse correctly."""
        result = split_top_level_choose("=CHOOSE(A1, B1, C1, D1)")
        assert result == ["A1", "B1", "C1", "D1"]

    def test_choose_with_trailing_addition_returns_none(self) -> None:
        """CHOOSE with trailing +E1 is NOT a top-level CHOOSE."""
        result = split_top_level_choose("=CHOOSE(A1, B1, C1)+E1")
        assert result is None


class TestSplitTopLevelSwitch:
    """Tests for split_top_level_switch function."""

    def test_simple_switch(self) -> None:
        """Basic SWITCH formula should parse correctly."""
        result = split_top_level_switch("=SWITCH(A1, 1, B1, 2, C1)")
        assert result == ["A1", "1", "B1", "2", "C1"]

    def test_switch_with_trailing_addition_returns_none(self) -> None:
        """SWITCH with trailing +D1 is NOT a top-level SWITCH."""
        result = split_top_level_switch("=SWITCH(A1, 1, B1)+D1")
        assert result is None

    def test_switch_with_xlfn_prefix(self) -> None:
        """SWITCH with _xlfn. prefix should parse correctly."""
        result = split_top_level_switch("=_xlfn.SWITCH(A1, 1, B1)")
        assert result == ["A1", "1", "B1"]
