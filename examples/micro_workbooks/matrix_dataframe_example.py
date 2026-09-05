#!/usr/bin/env python3
"""Pass a tidy pandas DataFrame to a generated matrix setter.

Demonstrates the **SeriesInput** DataFrame contract for a multi-key matrix binding:
one row per matrix cell, columns = binding ``key`` fields plus the measure concept
(``OBS_VALUE``). See also ``setter_dataframe_example.py`` for single-key series.

Run from the repo root::

    uv run python examples/micro_workbooks/matrix_dataframe_example.py
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

import pandas as pd

from excel_grapher.exporter import CodeGenerator
from excel_grapher.series_bindings.workflow import validate_bindings_workbook
from tests.fixtures.series_bindings.matrix_helpers import (
    MATRIX_EXPLICIT_BINDINGS,
    write_matrix_explicit_workbook,
)

REPO_ROOT = Path(__file__).resolve().parents[2]
BINDINGS = MATRIX_EXPLICIT_BINDINGS


def wide_matrix_to_tidy(df_wide: pd.DataFrame) -> pd.DataFrame:
    """Convert a wide indicator block (periods as columns) to tidy setter input."""
    return (
        df_wide.reset_index()
        .melt(id_vars=["INDICATOR"], var_name="TIME_PERIOD", value_name="OBS_VALUE")
        .astype({"TIME_PERIOD": int})
    )


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        workbook = tmp_path / "matrix_inputs.xlsx"
        write_matrix_explicit_workbook(workbook)

        result = validate_bindings_workbook(workbook, BINDINGS)
        with CodeGenerator(result["graph"]) as gen:
            code = gen.generate(
                result["targets"],
                series_bindings=result["bindings"],
                bindings_workbook=workbook,
            )
        namespace: dict[str, Any] = {}
        exec(code, namespace)
        setter = namespace["set_macro_matrix"]

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

        ctx = namespace["make_context"]()
        setter(ctx, updates)
        print("After partial DataFrame update:")
        print(f"  Inputs!C3 (GDP growth, 2025): {ctx.inputs['Inputs!C3']}")
        print(f"  Inputs!D5 (Debt, 2026): {ctx.inputs['Inputs!D5']}")
        print(f"  Inputs!B3 (GDP growth, 2024, unchanged): {ctx.inputs['Inputs!B3']}")
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

        for indicator in ["GDP growth"]:
            tidy_row = wide_matrix_to_tidy(df_wide.loc[[indicator]])
            print(f"Tidy rows for {indicator!r}:")
            print(tidy_row.to_string(index=False))
            print()
            ctx_row = namespace["make_context"]()
            setter(ctx_row, tidy_row, empty_measure="skip")
            print(f"After row update for {indicator!r}:")
            print(f"  Inputs!B3 (2024): {ctx_row.inputs['Inputs!B3']}")
            print(f"  Inputs!C3 (2025): {ctx_row.inputs['Inputs!C3']}")


if __name__ == "__main__":
    main()
