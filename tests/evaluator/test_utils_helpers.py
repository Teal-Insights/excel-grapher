"""Tests for utils/_helpers.py platform detection and path conversion."""

from __future__ import annotations

import os
from pathlib import Path
from unittest.mock import patch

import pytest

from tests.utils._helpers import (
    is_libreoffice_available,
    is_wsl,
    parse_cell_ref,
    wsl_path_to_windows_unc,
)


class TestIsWsl:
    """Tests for is_wsl() function."""

    def test_detects_wsl_from_release(self) -> None:
        """Test that WSL is detected from 'microsoft' in platform release."""
        with patch("platform.release", return_value="5.15.0-1-Microsoft"):
            assert is_wsl() is True

    def test_detects_wsl_lowercase(self) -> None:
        """Test case-insensitive WSL detection."""
        with patch("platform.release", return_value="5.15.0-1-microsoft-standard-WSL2"):
            assert is_wsl() is True

    def test_non_wsl_linux(self) -> None:
        """Test that regular Linux is not detected as WSL."""
        with patch("platform.release", return_value="6.8.0-90-generic"):
            assert is_wsl() is False

    def test_non_wsl_macos(self) -> None:
        """Test that macOS is not detected as WSL."""
        with patch("platform.release", return_value="23.5.0"):
            assert is_wsl() is False


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

    def test_quoted_sheet_name(self) -> None:
        """Test parsing a cell reference with quoted sheet name."""
        sheet, cell = parse_cell_ref("'Sheet Name'!$C$5")
        assert sheet == "Sheet Name"
        assert cell == "C5"

    def test_quoted_sheet_with_spaces(self) -> None:
        """Test parsing a cell reference with spaces in sheet name."""
        sheet, cell = parse_cell_ref("'My Data Sheet'!D14")
        assert sheet == "My Data Sheet"
        assert cell == "D14"

    def test_mixed_absolute_reference(self) -> None:
        """Test parsing a mixed absolute/relative reference."""
        sheet, cell = parse_cell_ref("Sales!$A2")
        assert sheet == "Sales"
        assert cell == "A2"

    def test_invalid_format_raises(self) -> None:
        """Test that invalid format raises ValueError."""
        with pytest.raises(ValueError, match="sheet-qualified"):
            parse_cell_ref("NoExclamationMark")


class TestIsLibreofficeAvailable:
    """Tests for is_libreoffice_available() function."""

    def test_available_when_soffice_found(self) -> None:
        """Test returns True when soffice is in PATH."""
        with patch("shutil.which", return_value="/usr/bin/soffice"):
            assert is_libreoffice_available() is True

    def test_unavailable_when_soffice_not_found(self) -> None:
        """Test returns False when soffice is not in PATH."""
        with patch("shutil.which", return_value=None):
            assert is_libreoffice_available() is False
