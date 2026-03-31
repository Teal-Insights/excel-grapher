"""Excel modification and recalculation utilities.

Supports multiple backends:
- WSL: PowerShell/COM to Windows Excel
- Windows/macOS: xlwings
- Linux: LibreOffice (fallback)
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from textwrap import dedent

from tests.utils._helpers import (
    check_libreoffice_version,
    is_libreoffice_available,
    is_wsl,
    parse_cell_ref,
)


class ExcelRecalculationError(Exception):
    """Raised when Excel recalculation fails."""

    pass


def _modify_and_recalculate_with_powershell(
    input_path: Path,
    output_path: Path,
    cell_modifications: dict[str, float],
) -> None:
    """Use Windows Excel via PowerShell/COM from WSL.

    Args:
        input_path: Path to the source Excel workbook.
        output_path: Path to save the modified workbook.
        cell_modifications: Dict mapping cell references to new values.

    Raises:
        ExcelRecalculationError: If PowerShell execution fails.
    """
    from tests.utils._helpers import wsl_path_to_windows_unc

    # Convert WSL paths to Windows UNC paths
    input_unc = wsl_path_to_windows_unc(input_path)
    output_unc = wsl_path_to_windows_unc(output_path)

    # Build cell modification PowerShell statements
    modifications = []
    for cell_ref, value in cell_modifications.items():
        sheet_name, cell_address = parse_cell_ref(cell_ref)
        # Escape single quotes in sheet names
        sheet_name_escaped = sheet_name.replace("'", "''")
        modifications.append(
            f'$wb.Worksheets.Item("{sheet_name_escaped}").Range("{cell_address}").Value2 = {value}'
        )

    modifications_script = "\n".join(modifications)

    script = dedent(f"""\
        $ErrorActionPreference = "Stop"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try {{
            $wb = $excel.Workbooks.Open("{input_unc}")

            # Modify cells
            {modifications_script}

            # Recalculate
            $wb.RefreshAll()
            foreach($ws in $wb.Worksheets) {{ $ws.Calculate() }}
            $excel.CalculateUntilAsyncQueriesDone()

            # Save as new file
            $wb.SaveAs("{output_unc}")
            $wb.Close($false)
        }} finally {{
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }}
    """)

    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command", script],
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        raise ExcelRecalculationError(
            f"PowerShell Excel recalculation failed:\n"
            f"stdout: {result.stdout}\n"
            f"stderr: {result.stderr}"
        )


def _modify_and_recalculate_with_xlwings(
    input_path: Path,
    output_path: Path,
    cell_modifications: dict[str, float],
) -> None:
    """Use xlwings for Windows/macOS Excel automation.

    Args:
        input_path: Path to the source Excel workbook.
        output_path: Path to save the modified workbook.
        cell_modifications: Dict mapping cell references to new values.

    Raises:
        ImportError: If xlwings is not installed.
        ExcelRecalculationError: If Excel automation fails.
    """
    import importlib
    from typing import Any, cast

    try:
        xw = importlib.import_module("xlwings")
        xw_any = cast(Any, xw)
    except ImportError as e:
        raise ImportError(
            "xlwings is required for Excel automation on Windows/macOS. "
            "Install with: pip install xlwings"
        ) from e

    app = xw_any.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(input_path))

        for cell_ref, value in cell_modifications.items():
            sheet_name, cell_address = parse_cell_ref(cell_ref)
            sheet = wb.sheets[sheet_name]
            sheet.range(cell_address).value = value

        # Recalculate
        wb.api.Calculate()

        # Save to output path
        wb.save(str(output_path))
        wb.close()
    except Exception as e:
        raise ExcelRecalculationError(f"xlwings Excel recalculation failed: {e}") from e
    finally:
        app.quit()


def _modify_and_recalculate_with_libreoffice(
    input_path: Path,
    output_path: Path,
    cell_modifications: dict[str, float],
    timeout: int = 120,
) -> None:
    """Use LibreOffice for headless Excel recalculation on Linux.

    This process:
    1. Modifies cells with fastpyxl (no recalculation)
    2. Uses LibreOffice to open and re-save the file, triggering recalculation

    Args:
        input_path: Path to the source Excel workbook.
        output_path: Path to save the modified workbook.
        cell_modifications: Dict mapping cell references to new values.
        timeout: Timeout in seconds for LibreOffice execution.

    Raises:
        ExcelRecalculationError: If LibreOffice recalculation fails.
    """
    check_libreoffice_version()

    from fastpyxl import load_workbook

    # Step 1: Copy and modify with fastpyxl
    shutil.copy2(input_path, output_path)

    wb = load_workbook(str(output_path), keep_vba=True)
    try:
        for cell_ref, value in cell_modifications.items():
            sheet_name, cell_address = parse_cell_ref(cell_ref)
            wb[sheet_name][cell_address] = value
        wb.save(str(output_path))
    finally:
        wb.close()

    # Step 2: Use LibreOffice to recalculate by converting to xlsx
    # LibreOffice will recalculate all formulas when it opens the file
    # We convert to xlsx and back to xlsm to force recalculation
    temp_xlsx = output_path.with_suffix(".xlsx")
    output_dir = output_path.parent

    try:
        # Convert xlsm to xlsx (triggers recalculation)
        cmd = [
            "soffice",
            "--headless",
            "--norestore",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_dir),
            str(output_path.resolve()),
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        if result.returncode != 0:
            raise ExcelRecalculationError(
                f"LibreOffice xlsx conversion failed:\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        # Check if the xlsx file was created
        if not temp_xlsx.exists():
            raise ExcelRecalculationError(
                f"LibreOffice did not create expected output file: {temp_xlsx}"
            )

        # Convert back to xlsm (to preserve macro structure if any)
        cmd = [
            "soffice",
            "--headless",
            "--norestore",
            "--convert-to",
            "xlsm",
            "--outdir",
            str(output_dir),
            str(temp_xlsx),
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        if result.returncode != 0:
            raise ExcelRecalculationError(
                f"LibreOffice xlsm conversion failed:\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

    except subprocess.TimeoutExpired as e:
        raise ExcelRecalculationError(
            f"LibreOffice recalculation timed out after {timeout}s"
        ) from e
    finally:
        # Clean up temp xlsx file
        temp_xlsx.unlink(missing_ok=True)


def modify_and_recalculate_workbook(
    input_path: Path,
    output_path: Path,
    cell_modifications: dict[str, float],
) -> None:
    """Modify cell values in an Excel workbook and recalculate all formulas.

    Auto-selects the appropriate backend based on the platform:
    - WSL: PowerShell/COM to Windows Excel
    - Windows/macOS: xlwings
    - Linux: LibreOffice (fallback)

    Args:
        input_path: Path to the source Excel workbook.
        output_path: Path to save the modified and recalculated workbook.
        cell_modifications: Dict mapping cell references to new values.
            Keys should be in format "'Sheet Name'!$A$1" or "Sheet!A1".
            Values should be numeric.

    Raises:
        ExcelRecalculationError: If recalculation fails.
        RuntimeError: If no suitable Excel automation backend is available.

    Example:
        >>> modify_and_recalculate_workbook(
        ...     input_path=Path("input.xlsx"),
        ...     output_path=Path("output.xlsx"),
        ...     cell_modifications={"'Data'!$A$1": 123.45, "Sheet1!B2": 67.89},
        ... )
    """
    if is_wsl():
        _modify_and_recalculate_with_powershell(input_path, output_path, cell_modifications)
    elif sys.platform in ("win32", "darwin"):
        _modify_and_recalculate_with_xlwings(input_path, output_path, cell_modifications)
    elif is_libreoffice_available():
        _modify_and_recalculate_with_libreoffice(input_path, output_path, cell_modifications)
    else:
        raise RuntimeError(
            "No suitable Excel automation backend available. "
            "On Linux, install LibreOffice: sudo apt install libreoffice"
        )
