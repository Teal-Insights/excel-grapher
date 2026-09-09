"""Shared axis constants for generated named-coordinate modules.

Every non-scalar series domain is built from `Axis` values. Identical axes
(same name, keys, and key type) are emitted once in `data.py` and referenced
by constant name from domains and from formula bodies that move along an
axis (`OFFSET`).
"""

from __future__ import annotations

import keyword
import re
from collections.abc import Iterable
from dataclasses import dataclass, field

from excel_grapher.exporter.export_runtime.tensor import Axis

_AXIS_KEY = tuple[str, tuple[str | int, ...], str]


def python_identifier(text: str) -> str:
    """Return a Python identifier derived from an authored name."""
    candidate = re.sub(r"\W", "_", text)
    if not candidate or candidate[0].isdigit():
        candidate = "_" + candidate
    while keyword.iskeyword(candidate):
        candidate += "_"
    return candidate


def _axis_key(axis: Axis) -> _AXIS_KEY:
    return (axis.name, axis.keys, axis.key_type.__name__)


@dataclass
class NamedAxes:
    """Deterministic constant names for the distinct axes of a catalog."""

    _names: dict[_AXIS_KEY, str] = field(default_factory=dict)
    _axes: dict[str, Axis] = field(default_factory=dict)

    @classmethod
    def plan(cls, axes: Iterable[Axis]) -> NamedAxes:
        """Assign `<FIELD>_AXIS` names in first-seen order, suffixing collisions."""
        planned = cls()
        for axis in axes:
            planned.register(axis)
        return planned

    def register(self, axis: Axis) -> str:
        """Return the constant name for `axis`, allocating one when new."""
        key = _axis_key(axis)
        existing = self._names.get(key)
        if existing is not None:
            return existing
        base = python_identifier(axis.name.upper()) + "_AXIS"
        name = base
        suffix = 2
        while name in self._axes:
            name = f"{base}_{suffix}"
            suffix += 1
        self._names[key] = name
        self._axes[name] = axis
        return name

    def constant(self, axis: Axis) -> str:
        """Return the constant name of a registered axis."""
        return self._names[_axis_key(axis)]

    def items(self) -> list[tuple[str, Axis]]:
        """Return `(constant name, axis)` pairs in allocation order."""
        return list(self._axes.items())
