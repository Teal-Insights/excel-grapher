"""Tests for utils/_helpers.py platform detection and path conversion."""

from __future__ import annotations

import os
from pathlib import Path
from unittest.mock import patch

import pytest

from tests.utils._helpers import is_wsl, parse_cell_ref, wsl_path_to_windows_unc


class TestIsWsl:
    """Tests for is_wsl() function."""

    @pytest.mark.parametrize(
        ("release", "expected"),
        [
            ("5.15.0-1-Microsoft", True),
            ("5.15.0-1-microsoft-standard-WSL2", True),
            ("6.8.0-90-generic", False),
            ("23.5.0", False),
        ],
    )
    def test_detects_wsl_from_release(self, release: str, expected: bool) -> None:
        with patch("platform.release", return_value=release):
            assert is_wsl() is expected


class TestWslPathToWindowsUnc:
    """Tests for wsl_path_to_windows_unc() function."""

    def test_converts_simple_path(self) -> None:
        """Test converting a simple WSL path to UNC."""
        with patch.dict(os.environ, {"WSL_DISTRO_NAME": "Ubuntu"}):
            path = Path("/home/user/file.xlsx")
            with patch.object(Path, "resolve", return_value=path):
                result = wsl_path_to_windows_unc(path)
                assert result == r"\\wsl.localhost\Ubuntu\home\user\file.xlsx"

    def test_raises_without_distro_name(self) -> None:
        """Test that RuntimeError is raised when WSL_DISTRO_NAME is not set."""
        with patch.dict(os.environ, {}, clear=True):
            # Make sure WSL_DISTRO_NAME is not set
            os.environ.pop("WSL_DISTRO_NAME", None)
            path = Path("/home/user/file.xlsx")
            with pytest.raises(RuntimeError, match="WSL_DISTRO_NAME"):
                wsl_path_to_windows_unc(path)


class TestParseCellRef:
    """Tests for parse_cell_ref() function."""

    def test_simple_reference(self) -> None:
        """Test parsing a simple cell reference."""
        sheet, cell = parse_cell_ref("Sheet1!A1")
        assert sheet == "Sheet1"
        assert cell == "A1"

    def test_absolute_reference(self) -> None:
        """Test parsing an absolute cell reference with $ signs."""
        sheet, cell = parse_cell_ref("Data!$B$10")
        assert sheet == "Data"
        assert cell == "B10"

    @pytest.mark.parametrize(
        ("ref", "sheet_name", "cell"),
        [
            ("'Sheet Name'!$C$5", "Sheet Name", "C5"),
            ("'My Data Sheet'!D14", "My Data Sheet", "D14"),
        ],
    )
    def test_quoted_sheet_name(self, ref: str, sheet_name: str, cell: str) -> None:
        parsed_sheet, parsed_cell = parse_cell_ref(ref)
        assert parsed_sheet == sheet_name
        assert parsed_cell == cell

    def test_mixed_absolute_reference(self) -> None:
        """Test parsing a mixed absolute/relative reference."""
        sheet, cell = parse_cell_ref("Sales!$A2")
        assert sheet == "Sales"
        assert cell == "A2"

    def test_invalid_format_raises(self) -> None:
        """Test that invalid format raises ValueError."""
        with pytest.raises(ValueError, match="sheet-qualified"):
            parse_cell_ref("NoExclamationMark")
