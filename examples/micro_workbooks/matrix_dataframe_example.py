#!/usr/bin/env python3
"""Pass a tidy pandas DataFrame as catalog-order input to a matrix compute.

Inverted-tree packages take sequences in catalog order (row-major over the
matrix keys), not ctx setters. Convert a tidy table (binding ``key`` columns
plus ``OBS_VALUE``) into that tuple, then call ``compute_*``.

See also ``setter_dataframe_example.py`` for single-key series.

Run from the repo root::

    uv run python examples/micro_workbooks/matrix_dataframe_example.py
"""

from __future__ import annotations

import importlib
import sys
import tempfile
from pathlib import Path
from typing import Any

import pandas as pd
from fastpyxl import load_workbook

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.fixtures.series_bindings.matrix_helpers import (
    MACRO_MATRIX_PERIODS,
    macro_matrix_structure,
    write_matrix_explicit_workbook,
)


def add_formula_output_block(path: Path) -> None:
    """Mirror Inputs!B3:D5 into F3:H5 with copied year headers on row 2."""
    workbook = load_workbook(path)
    worksheet = workbook["Inputs"]
    for col_offset, period in enumerate(MACRO_MATRIX_PERIODS):
        src = chr(ord("B") + col_offset)
        dest = chr(ord("F") + col_offset)
        worksheet[f"{dest}2"] = period
        for row in range(3, 6):
            worksheet[f"{dest}{row}"] = f"={src}{row}"
    workbook.save(path)


def bindings_document(*, workbook: str) -> dict[str, Any]:
    structure = macro_matrix_structure()
    return {
        "schema_version": "1.4.0",
        "workbook": workbook,
        "series": [
            {
                "id": "macro_matrix",
                "sheet": "Inputs",
                "data_range": "Inputs!B3:D5",
                "layout": "matrix",
                "input": {"setter": {"name": "set_macro_matrix"}},
                "structure": structure,
                "key": ["INDICATOR", "TIME_PERIOD"],
            },
            {
                "id": "macro_result",
                "sheet": "Inputs",
                "data_range": "Inputs!F3:H5",
                "layout": "matrix",
                "output": {"compute": {"name": "compute_macro_result"}},
                "structure": structure,
                "key": ["INDICATOR", "TIME_PERIOD"],
            },
        ],
    }


def wide_matrix_to_tidy(df_wide: pd.DataFrame) -> pd.DataFrame:
    """Convert a wide indicator block (periods as columns) to tidy input."""
    return (
        df_wide.reset_index()
        .melt(id_vars=["INDICATOR"], var_name="TIME_PERIOD", value_name="OBS_VALUE")
        .astype({"TIME_PERIOD": int})
    )


def apply_tidy_updates(
    default: tuple[object, ...],
    domain: tuple[object, ...],
    keys: tuple[str, ...],
    tidy: pd.DataFrame,
) -> tuple[object, ...]:
    """Overlay tidy OBS_VALUE rows onto a catalog-order default tuple."""
    values = list(default)
    for row in tidy.to_dict(orient="records"):
        point = tuple(row[key] for key in keys)
        values[domain.index(point)] = float(row["OBS_VALUE"])
    return tuple(values)


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        workbook = tmp_path / "matrix_inputs.xlsx"
        write_matrix_explicit_workbook(workbook)
        add_formula_output_block(workbook)
        bindings = validate_bindings_document(bindings_document(workbook=workbook.name))
        targets = all_series_targets(bindings, workbook=workbook)
        graph = create_dependency_graph(workbook, targets, load_values=True)
        with CodeGenerator(graph) as gen:
            modules = gen.generate_modules(
                series_bindings=bindings,
                bindings_workbook=workbook,
            )
        pkg_dir = tmp_path / "matrix_df_pkg"
        pkg_dir.mkdir()
        for filename, content in modules.items():
            (pkg_dir / filename).write_text(content, encoding="utf-8")
        sys.path.insert(0, str(tmp_path))
        pkg = importlib.import_module("matrix_df_pkg")
        domain = pkg.compute_macro_result.__domain__
        keys = pkg.compute_macro_result.__key__
        default = pkg.data.MACRO_MATRIX_DEFAULT

        # --- 1. Tidy DataFrame (partial update: two cells only) ---
        updates = pd.DataFrame(
            {
                "INDICATOR": ["GDP growth", "Debt"],
                "TIME_PERIOD": [2025, 2026],
                "OBS_VALUE": [9.9, 44.4],
            }
        )
        print("Tidy matrix input:")
        print(updates.to_string(index=False))
        print()

        overlay = apply_tidy_updates(default, domain, keys, updates)
        result = pkg.compute_macro_result(macro_matrix=overlay)
        records = pkg.as_records(pkg.compute_macro_result, result)
        by_key = {(row["INDICATOR"], row["TIME_PERIOD"]): row["OBS_VALUE"] for row in records}
        print("After partial DataFrame overlay:")
        print(f"  GDP growth 2025: {by_key[('GDP growth', 2025)]}")
        print(f"  Debt 2026: {by_key[('Debt', 2026)]}")
        print(f"  GDP growth 2024 (unchanged): {by_key[('GDP growth', 2024)]}")
        print()

        # --- 2. Row-by-row editing from a wide block ---
        df_wide = pd.DataFrame(
            {
                2024: [1.6, 3.0, 54.0],
                2025: [1.7, 2.8, 53.5],
                2026: [1.8, 2.6, 53.0],
            },
            index=["GDP growth", "Inflation", "Debt"],
        )
        df_wide.index.name = "INDICATOR"
        print("Wide matrix (one indicator row at a time):")
        print(df_wide.loc[["GDP growth"]].to_string())
        print()

        tidy_row = wide_matrix_to_tidy(df_wide.loc[["GDP growth"]])
        print("Tidy rows for 'GDP growth':")
        print(tidy_row.to_string(index=False))
        print()
        overlay = apply_tidy_updates(default, domain, keys, tidy_row)
        result = pkg.compute_macro_result(macro_matrix=overlay)
        records = pkg.as_records(pkg.compute_macro_result, result)
        by_key = {(row["INDICATOR"], row["TIME_PERIOD"]): row["OBS_VALUE"] for row in records}
        print("After row overlay for 'GDP growth':")
        print(f"  GDP growth 2024: {by_key[('GDP growth', 2024)]}")
        print(f"  GDP growth 2025: {by_key[('GDP growth', 2025)]}")


if __name__ == "__main__":
    main()
