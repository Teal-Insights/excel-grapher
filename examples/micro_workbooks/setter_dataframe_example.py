#!/usr/bin/env python3
"""Pass a tidy pandas DataFrame to a generated series setter.

Demonstrates the **SeriesInput** DataFrame contract for a single-key series binding:
one row per observation, columns = binding ``key`` fields plus the measure concept
(``OBS_VALUE``). Binding metadata (country, indicator, units) is omitted.

See also ``matrix_dataframe_example.py`` for multi-key matrix bindings.

Run from the repo root::

    uv run python examples/micro_workbooks/setter_dataframe_example.py
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

import pandas as pd
import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.series_bindings.workflow import validate_bindings_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
BINDINGS = REPO_ROOT / "tests/fixtures/series_bindings/borvelia_primary_balance.yaml"


def write_borvelia_workbook(path: Path) -> None:
    """Write a minimal workbook matching ``borvelia_primary_balance.yaml``."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):  # columns F..J
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))  # F5=-2, G5=-1, ..., J5=1
    wb.close()


def wide_row_to_tidy(df_wide: pd.DataFrame) -> pd.DataFrame:
    """Convert one wide indicator row (periods as columns) to tidy setter input."""
    return (
        df_wide.stack()
        .reset_index()
        .rename(columns={"level_1": "TIME_PERIOD", 0: "OBS_VALUE"})
        .drop(columns=["level_0"])[["TIME_PERIOD", "OBS_VALUE"]]
    )


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        workbook = tmp_path / "lic_inputs.xlsx"
        write_borvelia_workbook(workbook)

        result = validate_bindings_workbook(workbook, BINDINGS)
        with CodeGenerator(result["graph"]) as gen:
            code = gen.generate(
                result["targets"],
                series_bindings=result["bindings"],
                bindings_workbook=workbook,
            )
        namespace: dict[str, Any] = {}
        exec(code, namespace)
        setter = namespace["set_borvelia_primary_balance"]

        # --- 1. Tidy DataFrame (partial update: periods 4 and 5 only) ---
        updates = pd.DataFrame(
            {
                "TIME_PERIOD": [4, 5],
                "OBS_VALUE": [7.5, 8.0],
            }
        )
        print("Tidy input:")
        print(updates.to_string(index=False))
        print()

        ctx = namespace["make_context"]()
        setter(ctx, updates)
        print("After partial DataFrame update:")
        print(f"  Inputs!I5 (period 4): {ctx.inputs['Inputs!I5']}")
        print(f"  Inputs!J5 (period 5): {ctx.inputs['Inputs!J5']}")
        print(f"  Inputs!F5 (period 1, unchanged): {ctx.inputs['Inputs!F5']}")
        print()

        # --- 2. Wide spreadsheet row → tidy before calling the setter ---
        df_wide = pd.DataFrame(
            {1: [-1.0], 2: [-0.5], 3: [0.0], 4: [0.5], 5: [1.0]},
            index=["Primary balance (% of GDP)"],
        )
        print("Wide row:")
        print(df_wide.to_string(index=False))
        print()
        tidy = wide_row_to_tidy(df_wide)
        print("Wide row converted to tidy:")
        print(tidy.to_string(index=False))
        print()

        ctx2 = namespace["make_context"]()
        setter(ctx2, tidy)
        print("After wide→tidy update:")
        print(f"  Inputs!I5 (period 4): {ctx2.inputs['Inputs!I5']}")
        print(f"  Inputs!J5 (period 5): {ctx2.inputs['Inputs!J5']}")


if __name__ == "__main__":
    main()
