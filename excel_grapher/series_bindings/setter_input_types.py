"""Type aliases for generated and library setter input coercion."""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING, Literal, TypeAlias

from excel_grapher.series_bindings.records_types import Record, Records, Scalar

Layout: TypeAlias = Literal["scalar", "series", "matrix"]

if TYPE_CHECKING:
    import pandas as pd
    import polars as pl

    DataFrameInput: TypeAlias = pd.DataFrame | pl.DataFrame
else:
    DataFrameInput: TypeAlias = object

SeriesInput: TypeAlias = Records | Record | Sequence[Scalar] | DataFrameInput
SetterInput: TypeAlias = SeriesInput | Scalar

__all__ = [
    "DataFrameInput",
    "Layout",
    "SeriesInput",
    "SetterInput",
]
