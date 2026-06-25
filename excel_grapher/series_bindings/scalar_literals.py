"""Render series-binding scalar values as Python literals."""

from __future__ import annotations

from datetime import datetime

__all__ = ["py_scalar_literal"]


def py_scalar_literal(value: object, *, datetime_ref: str = "datetime.datetime") -> str:
    """Render a scalar value as a Python literal for generated binding code."""
    if value is None:
        return "None"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, datetime):
        rendered = repr(value)
        if datetime_ref != "datetime.datetime":
            rendered = rendered.replace("datetime.datetime", datetime_ref, 1)
        return rendered
    if isinstance(value, str):
        return repr(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return repr(value)
    if isinstance(value, float):
        return repr(value)
    return repr(value)
