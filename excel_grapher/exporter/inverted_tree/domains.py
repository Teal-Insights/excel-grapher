"""Key-domain tuples for inverted-tree `data.py` and compute metadata.

Field domains (`TIME_PERIOD_DOMAIN`, ...) are the catalog-order union of resolved
key values, one tuple per distinct field. Each `compute_*` / internals helper
publishes `__key__` and `__domain__` so callers index by key instead of column
count (#676). Generated modules attach that metadata with `@publish` (#766).
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from dataclasses import field as data_field
from datetime import datetime

from excel_grapher.exporter.inverted_tree.catalog import BoundSeries
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings.types import Scalar

_LOOP_VARS = {
    "TIME_PERIOD": "period",
}


def domain_const_name(field: str) -> str:
    """Return the `data.py` constant name for key field `field`."""
    if not field.isidentifier() or field.startswith("_"):
        raise InvertedTreeExportError(f"key field {field!r} cannot be emitted as a domain constant")
    return f"{field}_DOMAIN"


def series_domain_points(series: BoundSeries) -> tuple[object, ...]:
    """Return this series' domain in catalog member order.

    A one-key series yields scalars so `TIME_PERIOD_DOMAIN.index(2050)` works.
    A multi-key series yields tuples `(outer, ..., TIME_PERIOD)`. A keyless
    series yields `()` once per member.
    """
    fields = series.key_fields
    if not fields:
        return tuple(() for _ in series.cells)
    points: list[object] = []
    for point in series.domain:
        values = tuple(point[field] for field in fields)
        points.append(values[0] if len(fields) == 1 else values)
    return tuple(points)


def _contiguous_slice(full: tuple[object, ...], part: tuple[object, ...]) -> slice | None:
    """Return a slice such that `full[slice] == part`, or None."""
    if not part:
        return None
    if part == full:
        return slice(None)
    length = len(full)
    width = len(part)
    for start, first in enumerate(full):
        if first != part[0]:
            continue
        for step in range(1, length):
            last = start + (width - 1) * step
            if last >= length:
                break
            candidate = full[start : last + 1 : step]
            if candidate != part:
                continue
            stop = None if last + step >= length else last + 1
            start_arg = None if start == 0 else start
            step_arg = None if step == 1 else step
            return slice(start_arg, stop, step_arg)
    return None


def _slice_source(slc: slice) -> str:
    """Return the `[start:stop:step]` suffix for `slc`."""
    start, stop, step = slc.start, slc.stop, slc.step
    if step is None or step == 1:
        start_s = "" if start is None else str(start)
        stop_s = "" if stop is None else str(stop)
        return f"[{start_s}:{stop_s}]"
    start_s = "" if start is None else str(start)
    stop_s = "" if stop is None else str(stop)
    return f"[{start_s}:{stop_s}:{step}]"


_MAX_PRODUCT_SCAN = 250_000
_MAX_EXCEPTIONS = 8


_MISSING = object()


def uses_datetime_values(values: Sequence[object]) -> bool:
    """True when `values` contains a `datetime` (including nested tuples)."""
    for value in values:
        if isinstance(value, datetime):
            return True
        if isinstance(value, tuple) and uses_datetime_values(value):
            return True
    return False


@dataclass(frozen=True, slots=True)
class DomainEmitPlan:
    """Expressions and interned tuples for one catalog's key domains.

    `interned_source` maps `_DOMAIN_N` names to compact `data.py` right-hand
    sides. Names absent from the map are emitted as tuple literals.
    """

    field_domains: dict[str, tuple[Scalar, ...]]
    interned: tuple[tuple[str, tuple[object, ...]], ...]
    series_expr: dict[str, str]
    series_key: dict[str, tuple[str, ...]]
    scc_expr: dict[tuple[str, ...], str]
    scc_key: dict[tuple[str, ...], tuple[str, ...]]
    interned_source: dict[str, str] = data_field(default_factory=dict)

    def uses_data(self, series_id: str) -> bool:
        """True when this series' `__domain__` expression reads `data`."""
        return "data." in self.series_expr.get(series_id, "")

    def uses_data_scc(self, scc: tuple[str, ...]) -> bool:
        """True when this SCC's `__domain__` expression reads `data`."""
        return "data." in self.scc_expr.get(scc, "")

    @property
    def uses_datetime(self) -> bool:
        """True when a field or interned domain contains a datetime."""
        if any(uses_datetime_values(values) for values in self.field_domains.values()):
            return True
        return any(uses_datetime_values(values) for _, values in self.interned)

    def any_data_ref(self) -> bool:
        """True when any published `__domain__` expression reads `data`."""
        if any("data." in expr for expr in self.series_expr.values()):
            return True
        return any("data." in expr for expr in self.scc_expr.values())

    def sequence_expr(self, values: tuple[object, ...], *, field: str | None = None) -> str | None:
        """Return a `data.py` expression for `values`, or `None` if not interned.

        Prefers a field-domain constant (or slice) when `field` is known,
        then any interned subset, then any field domain that matches exactly.
        """
        candidates: list[tuple[str, tuple[object, ...]]] = []
        if field is not None and field in self.field_domains:
            candidates.append((domain_const_name(field), self.field_domains[field]))
        candidates.extend(self.interned)
        if field is None:
            candidates.extend(
                (domain_const_name(name), domain) for name, domain in self.field_domains.items()
            )
        seen: set[str] = set()
        for name, full in candidates:
            if name in seen:
                continue
            seen.add(name)
            slc = _contiguous_slice(full, values)
            if slc is None:
                continue
            ref = f"data.{name}"
            return ref if slc == slice(None) else f"{ref}{_slice_source(slc)}"
        return None
