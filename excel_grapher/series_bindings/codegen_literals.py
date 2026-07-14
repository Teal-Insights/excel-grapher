"""Python literal emission for generated series-binding code."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from datetime import datetime

from excel_grapher.series_bindings.scalar_literals import py_scalar_literal
from excel_grapher.series_bindings.types import SeriesResolution

__all__ = [
    "emit_compute_preamble_lines",
    "emit_setter_type_alias_lines",
    "py_scalar_literal",
    "python_annotation_for_dtype",
    "resolution_includes_datetime",
    "resolutions_include_datetime",
    "setter_input_annotation",
    "values_include_datetime",
]


def python_annotation_for_dtype(dtype: str | None) -> str | None:
    """Return a Python type-annotation fragment for a binding dtype.

    Returns:
        Annotation text such as `float` or `int | float`, or `None` when the dtype
        is missing/`auto`/unknown.
    """
    if dtype is None or dtype == "auto":
        return None
    mapping = {
        "float": "float",
        "number": "int | float",
        "int": "int",
        "bool": "bool",
        "string": "str",
        "datetime": "datetime",
    }
    return mapping.get(dtype)


def setter_input_annotation(
    *,
    layout: str,
    measure_dtype: str | None,
    scalar_shorthand: bool,
) -> str:
    """Return the parameter annotation for a generated setter input.

    Scalar shorthand keeps a records union so dict/list inputs remain valid while
    exposing the measure dtype on the bare-value arm. Series/matrix narrow the
    positional measure sequence when a dtype is known.
    """
    measure_type = python_annotation_for_dtype(measure_dtype)
    if scalar_shorthand:
        if measure_type is None:
            return "Records | Record | Scalar"
        return f"Records | Record | {measure_type}"
    if measure_type is None:
        return "SeriesInput"
    if layout == "matrix":
        # Positional 1D measure iterables are not accepted for matrix setters.
        return "Records | Record | DataFrameInput"
    return f"Records | Record | Sequence[{measure_type}] | DataFrameInput"


def values_include_datetime(*containers: Mapping[str, object] | None) -> bool:
    """Return True when any mapped value is a ``datetime``."""
    for container in containers:
        if container is None:
            continue
        for value in container.values():
            if isinstance(value, datetime):
                return True
    return False


def resolution_includes_datetime(resolved: SeriesResolution) -> bool:
    """Return True when a resolution emits datetime scalars in keys or records."""
    for leaf in resolved["leaves"]:
        if values_include_datetime(leaf["key"], leaf["coordinates"], leaf["record"]):
            return True
    return False


def resolutions_include_datetime(resolutions: Iterable[SeriesResolution]) -> bool:
    """Return True when any resolution in ``resolutions`` uses datetime scalars."""
    return any(resolution_includes_datetime(resolved) for resolved in resolutions)


def emit_setter_type_alias_lines(*, include_datetime: bool) -> list[str]:
    """Emit shared type aliases for generated setter blocks."""
    sequence_import = [
        "from collections.abc import Sequence",
        "from typing import TYPE_CHECKING, TypeAlias",
        "",
    ]
    scalar_type = (
        "Scalar: TypeAlias = str | int | float | bool | datetime | None"
        if include_datetime
        else "Scalar: TypeAlias = str | int | float | bool | None"
    )
    return sequence_import + [
        scalar_type,
        "Record: TypeAlias = dict[str, object]",
        "Records: TypeAlias = list[Record]",
        "",
        "if TYPE_CHECKING:",
        "    import pandas as pd",
        "    import polars as pl",
        "",
        "    DataFrameInput: TypeAlias = pd.DataFrame | pl.DataFrame",
        "else:",
        "    DataFrameInput: TypeAlias = object",
        "",
        "SeriesInput: TypeAlias = Records | Record | Sequence[Scalar] | DataFrameInput",
        "",
    ]


def emit_compute_preamble_lines(*, include_datetime: bool) -> list[str]:
    """Emit shared type aliases (and optional datetime import) for compute blocks."""
    lines: list[str] = []
    if include_datetime:
        lines.extend(["import datetime", ""])
    lines.extend(
        [
            "Record = dict[str, object]",
            "Records = list[Record]",
            "",
        ]
    )
    return lines
