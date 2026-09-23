"""Shared axis constants for generated named-coordinate modules.

Every non-scalar series domain is built from `Axis` values. Axes that share a
concept and key type intern to one master whose keys are the union of
observed keys, in a stable order that preserves every subsequence. A series
whose keys are not a subsequence of that master (ragged Excel order) keeps a
cloned `Axis`. Formula bodies and domains refer to the emitted constant by
name (`OFFSET`, `span`, `coordinate_runs`).
"""

from __future__ import annotations

import keyword
import re
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field

from excel_grapher.exporter.export_runtime.tensor import Axis

_AXIS_KEY = tuple[str, tuple[str | int, ...], str]
_KeyTuple = tuple[str | int, ...]


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


def is_subsequence(keys: Sequence[object], master: Sequence[object]) -> bool:
    """Return whether `keys` appear in order inside `master`."""
    iterator = iter(master)
    return all(key in iterator for key in keys)


def _master_keys(observed: Sequence[_KeyTuple]) -> _KeyTuple:
    """Choose a union order that maximizes subsequence sharing.

    Candidates are first-seen union, sorted unique keys when comparable, and
    the longest observed tuple with leftover keys appended. Sorted order wins
    ties so consecutive integer ranges share one timeline.
    """
    first_seen: list[str | int] = []
    seen: set[str | int] = set()
    for keys in observed:
        for key in keys:
            if key not in seen:
                seen.add(key)
                first_seen.append(key)
    first_seen_keys = tuple(first_seen)
    candidates: list[_KeyTuple] = [first_seen_keys]
    sorted_keys: _KeyTuple | None
    try:
        sorted_keys = tuple(sorted(first_seen))
    except TypeError:
        sorted_keys = None
    else:
        if sorted_keys not in candidates:
            candidates.append(sorted_keys)
    longest = max(observed, key=len)
    filled = tuple(longest) + tuple(key for key in first_seen if key not in set(longest))
    if filled not in candidates:
        candidates.append(filled)

    def rank(master: _KeyTuple) -> tuple[int, bool]:
        matches = sum(1 for keys in observed if is_subsequence(keys, master))
        return (matches, master == sorted_keys)

    return max(candidates, key=rank)


def layout_keys_source(axis: Axis, named: NamedAxes, *, module: str = "") -> str:
    """Source for this series' keys: Axis constant, `span`, or a key tuple.

    `module` is `data.` when the caller is outside `data.py`.
    """
    name = f"{module}{named.constant(axis)}"
    emitted = named.emitted(axis)
    keys = axis.keys
    if keys == emitted.keys:
        return name
    if not keys:
        return "()"
    if len(keys) == 1:
        return f"({keys[0]!r},)"
    try:
        start = emitted.keys.index(keys[0])
        stop = emitted.keys.index(keys[-1])
    except ValueError:
        return repr(keys)
    if keys == emitted.keys[start : stop + 1]:
        return f"span({name}, {keys[0]!r}, {keys[-1]!r})"
    return repr(keys)


@dataclass
class NamedAxes:
    """Deterministic constant names for the distinct axes of a catalog."""

    _names: dict[_AXIS_KEY, str] = field(default_factory=dict)
    _axes: dict[str, Axis] = field(default_factory=dict)

    @classmethod
    def plan(cls, axes: Iterable[Axis]) -> NamedAxes:
        """Assign `<FIELD>_AXIS` names, unioning subsequence key sets per concept."""
        observed = list(axes)
        planned = cls()
        groups: dict[tuple[str, str], list[Axis]] = {}
        order: list[tuple[str, str]] = []
        for axis in observed:
            group = (axis.name, axis.key_type.__name__)
            if group not in groups:
                order.append(group)
                groups[group] = []
            groups[group].append(axis)
        for group in order:
            members = groups[group]
            master_keys = _master_keys([axis.keys for axis in members])
            share = [axis for axis in members if is_subsequence(axis.keys, master_keys)]
            if share:
                master = Axis(members[0].name, master_keys, members[0].key_type)
                planned.register(master)
                master_name = planned.constant(master)
                for axis in share:
                    planned._names[_axis_key(axis)] = master_name
            for axis in members:
                if not is_subsequence(axis.keys, master_keys):
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

    def contains(self, axis: Axis) -> bool:
        """Return whether `plan` or `register` assigned `axis` a constant."""
        return _axis_key(axis) in self._names

    def constant(self, axis: Axis) -> str:
        """Return the constant name of a registered axis."""
        return self._names[_axis_key(axis)]

    def emitted(self, axis: Axis) -> Axis:
        """Return the `Axis` object emitted for `axis`, after unioning."""
        return self._axes[self.constant(axis)]

    def items(self) -> list[tuple[str, Axis]]:
        """Return `(constant name, axis)` pairs in allocation order."""
        return list(self._axes.items())
