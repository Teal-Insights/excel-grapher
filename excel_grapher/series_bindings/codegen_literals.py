"""Python literal emission for generated series-binding code."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from datetime import datetime

from excel_grapher.series_bindings.types import SeriesResolution

__all__ = [
    "emit_compute_preamble_lines",
    "emit_setter_type_alias_lines",
    "py_scalar_literal",
    "resolution_includes_datetime",
    "resolutions_include_datetime",
    "values_include_datetime",
]


def py_scalar_literal(value: object) -> str:
    """Render a scalar value as a Python literal for generated binding code."""
    if value is None:
        return "None"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, datetime):
        return repr(value)
    if isinstance(value, str):
        return repr(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return repr(value)
    if isinstance(value, float):
        return repr(value)
    return repr(value)


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
    if include_datetime:
        return [
            "import datetime",
            "",
            "Scalar = str | int | float | bool | datetime.datetime | None",
            "Record = dict[str, object]",
            "Records = list[Record]",
            "",
        ]
    return [
        "Scalar = str | int | float | bool | None",
        "Record = dict[str, object]",
        "Records = list[Record]",
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
