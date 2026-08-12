from __future__ import annotations

from pathlib import Path

import fastpyxl
import fastpyxl.utils.cell


def discover_formula_cells_in_rows(
    wb_path: Path,
    sheet_name: str,
    rows: list[int],
    *,
    wb_formulas: fastpyxl.Workbook | None = None,
    wb_values: fastpyxl.Workbook | None = None,
) -> list[str]:
    """Scan specified rows and return sheet-qualified addresses for formula cells.

    Only includes cells that contain formulas (start with '=') and whose cached
    calculated value is numeric.

    Pass a pre-opened *wb_formulas* loaded with `keep_formula_cache=True` to
    avoid repeated `load_workbook` calls when scanning multiple sheets.
    *wb_values* is accepted for backward compatibility when the formulas
    workbook was not loaded with formula caches; prefer a single dual-load
    workbook instead.
    """
    owned = wb_formulas is None
    if wb_formulas is None:
        wb_formulas = fastpyxl.load_workbook(
            wb_path,
            data_only=False,
            read_only=True,
            keep_vba=True,
            keep_formula_cache=True,
        )
    use_side_cache = bool(getattr(wb_formulas, "keep_formula_cache", False))
    owned_values = False
    values_wb: fastpyxl.Workbook | None = None
    if not use_side_cache:
        if wb_values is None:
            values_wb = fastpyxl.load_workbook(
                wb_path, data_only=True, read_only=True, keep_vba=True
            )
            owned_values = owned
        else:
            values_wb = wb_values
    try:
        if sheet_name not in wb_formulas.sheetnames:
            print(f"  Warning: Sheet '{sheet_name}' not found")
            return []
        if values_wb is not None and sheet_name not in values_wb.sheetnames:
            print(f"  Warning: Sheet '{sheet_name}' not found")
            return []

        ws_formulas = wb_formulas[sheet_name]
        ws_values = None if values_wb is None else values_wb[sheet_name]
        targets: list[str] = []

        for row in rows:
            max_col = ws_formulas.max_column or 1
            for col_idx in range(1, max_col + 1):
                cell_formula = ws_formulas.cell(row=row, column=col_idx)
                if isinstance(cell_formula.value, str) and cell_formula.value.startswith("="):
                    if ws_values is None:
                        cached_value = cell_formula.cached_value
                    else:
                        cached_value = ws_values.cell(row=row, column=col_idx).value
                    if not isinstance(cached_value, (int, float)) or isinstance(cached_value, bool):
                        continue
                    col_letter = fastpyxl.utils.cell.get_column_letter(col_idx)
                    targets.append(f"{sheet_name}!{col_letter}{row}")

        return targets
    finally:
        if owned:
            wb_formulas.close()
        if owned_values and values_wb is not None:
            values_wb.close()
