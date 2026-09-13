#!/usr/bin/env python3
"""Pass tidy DataFrame records to a named-coordinate generated function.

Overlay rows by their declared `TIME_PERIOD` keys and publish a new immutable
input tensor. Run with `uv run python examples/micro_workbooks/setter_dataframe_example.py`.
"""

from __future__ import annotations

import importlib
import sys
import tempfile
from pathlib import Path
from typing import Any

import pandas as pd
import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets


def write_borvelia_workbook(path: Path) -> None:
    """Write input values on F5:J5 and identity formulas on F6:J6."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):  # columns F..J
        letter = "FGHIJ"[col - 5]
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))  # F5=-2, G5=-1, ..., J5=1
        ws.write_formula(5, col, f"={letter}5")
    wb.close()


def _structure() -> dict[str, Any]:
    return {
        "measure": {
            "concept": "OBS_VALUE",
            "dtype": "float",
            "bind": {"kind": "data_cell", "read": "float"},
        },
        "dimensions": [
            {
                "concept": "TIME_PERIOD",
                "role": "key",
                "scope": "cell",
                "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
            }
        ],
    }


def bindings_document(*, workbook: str) -> dict[str, Any]:
    return {
        "schema_version": "1.3.0",
        "workbook": workbook,
        "series": [
            {
                "id": "borvelia_primary_balance",
                "sheet": "Inputs",
                "data_range": "Inputs!F5:J5",
                "layout": "series",
                "input": {"setter": {"name": "set_borvelia_primary_balance"}},
                "structure": _structure(),
                "key": ["TIME_PERIOD"],
            },
            {
                "id": "borvelia_primary_balance_out",
                "sheet": "Inputs",
                "data_range": "Inputs!F6:J6",
                "layout": "series",
                "output": {"compute": {"name": "compute_borvelia_primary_balance_out"}},
                "structure": _structure(),
                "key": ["TIME_PERIOD"],
            },
        ],
    }


def wide_row_to_tidy(df_wide: pd.DataFrame) -> pd.DataFrame:
    """Convert one wide indicator row (periods as columns) to tidy input."""
    return (
        df_wide.stack()
        .reset_index()
        .rename(columns={"level_1": "TIME_PERIOD", 0: "OBS_VALUE"})
        .drop(columns=["level_0"])[["TIME_PERIOD", "OBS_VALUE"]]
    )


def apply_tidy_updates(default: Any, tidy: pd.DataFrame) -> Any:
    """Overlay unique, valid period records onto an immutable input tensor."""
    values = dict(default.items())
    seen = set()
    for period, value in zip(tidy["TIME_PERIOD"], tidy["OBS_VALUE"], strict=True):
        coordinate = (period,)
        default.domain.require(coordinate)
        if coordinate in seen:
            raise ValueError(f"duplicate update for {coordinate!r}")
        seen.add(coordinate)
        values[coordinate] = float(value)
    return type(default).from_records(domain=default.domain, records=values.items())


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        workbook = tmp_path / "lic_inputs.xlsx"
        write_borvelia_workbook(workbook)
        bindings = validate_bindings_document(bindings_document(workbook=workbook.name))
        targets = all_series_targets(bindings, workbook=workbook)
        graph = create_dependency_graph(workbook, targets, load_values=True)
        with CodeGenerator(graph) as gen:
            modules = gen.generate_modules(
                series_bindings=bindings,
                bindings_workbook=workbook,
            )
        pkg_dir = tmp_path / "setter_df_pkg"
        pkg_dir.mkdir()
        for filename, content in modules.items():
            (pkg_dir / filename).write_text(content, encoding="utf-8")
        sys.path.insert(0, str(tmp_path))
        pkg = importlib.import_module("setter_df_pkg")
        default = pkg.data.BORVELIA_PRIMARY_BALANCE_DEFAULT

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

        overlay = apply_tidy_updates(default, updates)
        result = pkg.compute_borvelia_primary_balance_out(borvelia_primary_balance=overlay)
        records = pkg.as_records(pkg.compute_borvelia_primary_balance_out, result)
        by_period = {row["TIME_PERIOD"]: row["OBS_VALUE"] for row in records}
        print("After partial DataFrame overlay:")
        print(f"  period 4: {by_period[4]}")
        print(f"  period 5: {by_period[5]}")
        print(f"  period 1 (unchanged): {by_period[1]}")
        print()

        # --- 2. Wide spreadsheet row → tidy → named-coordinate tensor ---
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

        overlay = apply_tidy_updates(default, tidy)
        result = pkg.compute_borvelia_primary_balance_out(borvelia_primary_balance=overlay)
        records = pkg.as_records(pkg.compute_borvelia_primary_balance_out, result)
        by_period = {row["TIME_PERIOD"]: row["OBS_VALUE"] for row in records}
        print("After wide→tidy overlay:")
        print(f"  period 4: {by_period[4]}")
        print(f"  period 5: {by_period[5]}")


if __name__ == "__main__":
    main()
