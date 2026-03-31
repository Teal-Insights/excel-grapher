"""Platform detection and path conversion helpers."""

from __future__ import annotations

import os
import platform
from pathlib import Path


def is_wsl() -> bool:
    """Detect whether we are running inside WSL (Windows Subsystem for Linux).

    Returns:
        True if running in WSL, False otherwise.
    """
    return "microsoft" in platform.release().lower()


def wsl_path_to_windows_unc(path: Path) -> str:
    """Convert a WSL path to a Windows UNC path for Excel COM automation.

    This is needed when using Windows Excel from WSL via PowerShell/COM.

    Args:
        path: A WSL filesystem path.

    Returns:
        Windows UNC path (e.g., \\\\wsl.localhost\\Ubuntu\\home\\user\\file.xlsx)

    Raises:
        RuntimeError: If WSL_DISTRO_NAME environment variable is not set.
    """
    distro = os.environ.get("WSL_DISTRO_NAME")
    if not distro:
        raise RuntimeError(
            "WSL_DISTRO_NAME environment variable not set. "
            "This function can only be used inside WSL."
        )
    abs_path = path.resolve()
    rel = str(abs_path).lstrip("/")
    return r"\\wsl.localhost\{}\{}".format(distro, rel.replace("/", "\\"))


def is_libreoffice_available() -> bool:
    """Check if LibreOffice is available on the system.

    Returns:
        True if soffice command is available, False otherwise.
    """
    import shutil

    return shutil.which("soffice") is not None


_MIN_LO_VERSION = (25, 8)


def check_libreoffice_version() -> None:
    """Assert that LibreOffice 25.8+ is installed.

    Raises:
        RuntimeError: If soffice is missing or its version is too old.
    """
    import re
    import subprocess

    try:
        result = subprocess.run(
            ["soffice", "--version"],
            capture_output=True,
            text=True,
            timeout=10,
        )
    except FileNotFoundError:
        raise RuntimeError("soffice not found — install LibreOffice 25.8+") from None

    m = re.search(r"(\d+)\.(\d+)", result.stdout)
    if not m:
        raise RuntimeError(f"Could not parse LibreOffice version from: {result.stdout.strip()}")

    major, minor = int(m.group(1)), int(m.group(2))
    if (major, minor) < _MIN_LO_VERSION:
        raise RuntimeError(f"LibreOffice {major}.{minor} is too old — 25.8+ is required")


def parse_cell_ref(cell_ref: str) -> tuple[str, str]:
    """Parse a cell reference into sheet name and cell address.

    Args:
        cell_ref: Cell reference in format "'Sheet Name'!$A$1" or "Sheet!A1"

    Returns:
        Tuple of (sheet_name, cell_address) with $ signs stripped from address.

    Examples:
        >>> parse_cell_ref("'Sheet 1'!$A$1")
        ('Sheet 1', 'A1')
        >>> parse_cell_ref("Data!B2")
        ('Data', 'B2')
    """
    from excel_grapher.evaluator.name_utils import parse_address

    sheet_name, cell_part = parse_address(cell_ref)
    # Strip $ signs from cell address (absolute reference markers)
    cell_address = cell_part.replace("$", "")
    return sheet_name, cell_address
