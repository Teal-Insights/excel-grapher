"""Tiny DSA dynamic-ref constraints (vendored from tiny-dsa-extraction-pipeline).

Used by the Pass-1 canary graph build. Keep in sync with
``workbook_config.CONSTRAINTS`` in the sibling pipeline when that vintage changes.
"""

from __future__ import annotations

from typing import Annotated, Literal

from excel_grapher.core.cell_types import Between, RealBetween

_cols = ("C", "D", "E", "F", "G")

CONSTRAINTS: dict[str, object] = {
    "Inputs!A10": Literal["Borvelia"],
    "Inputs!A11": Literal["Litellia"],
    "Inputs!A12": Literal["Aurelium"],
    "Inputs!B22": Literal[1, 2, 3],
    "Inputs!B5": Literal["Borvelia", "Litellia", "Aurelium"],
    "Engine!C5": Literal[1],
    "Engine!D5": Literal[2],
    "Engine!E5": Literal[3],
    "Engine!F5": Literal[4],
    "Engine!G5": Literal[5],
    "Inputs!B10": Annotated[float, RealBetween(0.0, 200.0)],
    "Inputs!B11": Annotated[float, RealBetween(0.0, 200.0)],
    "Inputs!B12": Annotated[float, RealBetween(0.0, 200.0)],
    "Inputs!B21": Annotated[int, Between(1, 5)],
    "Inputs!B26": Annotated[float, RealBetween(-30.0, 30.0)],
    "Inputs!C26": Annotated[float, RealBetween(-30.0, 30.0)],
    "Inputs!D26": Annotated[float, RealBetween(-30.0, 30.0)],
    **{f"Inputs!{c}16": Annotated[float, RealBetween(-10.0, 15.0)] for c in _cols},
    **{f"Inputs!{c}17": Annotated[float, RealBetween(0.0, 20.0)] for c in _cols},
    **{f"Inputs!{c}18": Annotated[float, RealBetween(-15.0, 15.0)] for c in _cols},
}

BLANK_RANGES: tuple[str, ...] = ()
