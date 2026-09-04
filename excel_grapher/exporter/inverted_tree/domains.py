"""Key-domain tuples for inverted-tree `data.py` and compute metadata.

Field domains (`TIME_PERIOD_DOMAIN`, ...) are the catalog-order union of resolved
key values, one tuple per distinct field. Each `compute_*` / internals helper
publishes `__key__` and `__domain__` so callers index by key instead of column
count (#676).
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from itertools import product

from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog
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


def collect_field_domains(catalog: SeriesCatalog) -> dict[str, tuple[Scalar, ...]]:
    """Return unique key-field values in catalog order (first seen)."""
    domains: dict[str, tuple[Scalar, ...]] = {}
    seen: dict[str, set[Scalar]] = {}
    for series_id in catalog.order:
        series = catalog.get(series_id)
        for point in series.domain:
            mapping = point.as_mapping()
            for field in series.key_fields:
                if field not in mapping:
                    continue
                value = mapping[field]
                bucket = seen.setdefault(field, set())
                if value in bucket:
                    continue
                bucket.add(value)
                domains[field] = (*domains.get(field, ()), value)
    return domains


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


def _loop_var(field: str, used: set[str]) -> str:
    if field in _LOOP_VARS:
        candidate = _LOOP_VARS[field]
    else:
        candidate = field.rsplit("_", 1)[-1].lower()
        if not candidate.isidentifier():
            candidate = "key"
    base = candidate
    index = 2
    while candidate in used:
        candidate = f"{base}{index}"
        index += 1
    used.add(candidate)
    return candidate


def _product_comprehension(fields: tuple[str, ...], field_exprs: Sequence[str]) -> str:
    used: set[str] = set()
    names = [_loop_var(field, used) for field in fields]
    tuple_body = ", ".join(names)
    gens = " ".join(f"for {name} in {expr}" for name, expr in zip(names, field_exprs, strict=True))
    return f"tuple(({tuple_body}) {gens})"


def _field_values_from_points(
    fields: tuple[str, ...], points: tuple[object, ...]
) -> list[tuple[object, ...]]:
    per_field: list[list[object]] = [[] for _ in fields]
    seen: list[set[object]] = [set() for _ in fields]
    for point in points:
        if not isinstance(point, tuple) or len(point) != len(fields):
            raise InvertedTreeExportError(
                f"multi-key domain point {point!r} does not match key {fields!r}"
            )
        for index, value in enumerate(point):
            if value in seen[index]:
                continue
            seen[index].add(value)
            per_field[index].append(value)
    return [tuple(values) for values in per_field]


def _scalar_type_name(values: Sequence[object]) -> str:
    if not values:
        return "object"
    if all(isinstance(value, bool) for value in values):
        return "bool"
    if all(isinstance(value, int) and not isinstance(value, bool) for value in values):
        return "int"
    if all(isinstance(value, int | float) and not isinstance(value, bool) for value in values):
        return "float"
    if all(isinstance(value, str) for value in values):
        return "str"
    if all(isinstance(value, datetime) for value in values):
        return "datetime"
    return "object"


def domain_annotation(values: Sequence[object]) -> str:
    """Return a typing annotation for a domain tuple of `values`."""
    if values and isinstance(values[0], tuple):
        inner = ", ".join(_scalar_type_name((item,)) for item in values[0])
        if len(values[0]) == 1:
            inner += ","
        return f"tuple[tuple[{inner}], ...]"
    return f"tuple[{_scalar_type_name(values)}, ...]"


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
    """Expressions and interned tuples for one catalog's key domains."""

    field_domains: dict[str, tuple[Scalar, ...]]
    interned: tuple[tuple[str, tuple[object, ...]], ...]
    series_expr: dict[str, str]
    series_key: dict[str, tuple[str, ...]]
    scc_expr: dict[tuple[str, ...], str]
    scc_key: dict[tuple[str, ...], tuple[str, ...]]

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


class _Planner:
    """Build domain expressions, interning tuples that are not slices/products."""

    def __init__(self, field_domains: dict[str, tuple[Scalar, ...]]) -> None:
        self.field_domains = field_domains
        self.interned: list[tuple[str, tuple[object, ...]]] = []

    def intern(self, points: tuple[object, ...]) -> str:
        for name, values in self.interned:
            if values == points:
                return f"data.{name}"
        name = f"_DOMAIN_{len(self.interned)}"
        self.interned.append((name, points))
        return f"data.{name}"

    def field_ref(self, field: str, values: tuple[object, ...]) -> str | None:
        full = self.field_domains.get(field)
        if full is None:
            return None
        slc = _contiguous_slice(full, values)
        if slc is None:
            return None
        name = f"data.{domain_const_name(field)}"
        if slc == slice(None):
            return name
        return f"{name}{_slice_source(slc)}"

    def expr_for(self, keys: tuple[str, ...], points: tuple[object, ...]) -> str:
        if not keys:
            if points == ((),):
                return "((),)"
            if points and all(item == () for item in points):
                return f"((),) * {len(points)}"
            return self.intern(points)
        if not points:
            return "()"
        if len(keys) == 1:
            scalars = tuple(points)
            ref = self.field_ref(keys[0], scalars)
            return ref if ref is not None else self.intern(scalars)
        per_field = _field_values_from_points(keys, points)
        generated = tuple(product(*per_field))
        if generated == points:
            refs: list[str] = []
            for field, values in zip(keys, per_field, strict=True):
                ref = self.field_ref(field, values)
                if ref is None:
                    break
                refs.append(ref)
            else:
                return _product_comprehension(keys, refs)
        return self.intern(points)


def _ordered_unique_sccs(
    catalog: SeriesCatalog, scc_map: Mapping[str, tuple[str, ...]] | None
) -> list[tuple[str, ...]]:
    if not scc_map:
        return []
    seen: set[tuple[str, ...]] = set()
    ordered: list[tuple[str, ...]] = []
    for series_id in catalog.order:
        scc = scc_map.get(series_id)
        if scc is None or len(scc) < 2 or scc in seen:
            continue
        seen.add(scc)
        ordered.append(scc)
    return ordered


def _scc_key_and_points(
    scc: tuple[str, ...], catalog: SeriesCatalog
) -> tuple[tuple[str, ...], tuple[object, ...]]:
    members = [catalog.get(series_id) for series_id in catalog.order if series_id in set(scc)]
    if not members:
        members = [catalog.get(series_id) for series_id in scc]
    keys = members[0].key_fields
    if any(member.key_fields != keys for member in members):
        primary = max(members, key=lambda item: len(item.cells))
        return primary.key_fields, series_domain_points(primary)
    seen: set[object] = set()
    points: list[object] = []
    for member in members:
        for point in series_domain_points(member):
            if point in seen:
                continue
            seen.add(point)
            points.append(point)
    return keys, tuple(points)


def plan_domain_emission(
    catalog: SeriesCatalog,
    scc_map: Mapping[str, tuple[str, ...]] | None = None,
) -> DomainEmitPlan:
    """Plan field-domain constants and per-series `__domain__` expressions."""
    field_domains = collect_field_domains(catalog)
    reserved = {domain_const_name(field) for field in field_domains}
    for series in catalog.constant_series():
        if series.series_id.upper() in reserved:
            raise InvertedTreeExportError(
                f"constant series {series.series_id!r} collides with key domain "
                f"{series.series_id.upper()}"
            )
    for series in catalog.input_series():
        default_name = f"{series.series_id.upper()}_DEFAULT"
        if default_name in reserved:
            raise InvertedTreeExportError(
                f"input series {series.series_id!r} default {default_name} collides "
                "with a key domain constant"
            )
    planner = _Planner(field_domains)
    series_expr: dict[str, str] = {}
    series_key: dict[str, tuple[str, ...]] = {}
    for series_id in catalog.order:
        series = catalog.get(series_id)
        points = series_domain_points(series)
        series_key[series_id] = series.key_fields
        series_expr[series_id] = planner.expr_for(series.key_fields, points)
    scc_expr: dict[tuple[str, ...], str] = {}
    scc_key: dict[tuple[str, ...], tuple[str, ...]] = {}
    for scc in _ordered_unique_sccs(catalog, scc_map):
        keys, points = _scc_key_and_points(scc, catalog)
        scc_key[scc] = keys
        scc_expr[scc] = planner.expr_for(keys, points)
    return DomainEmitPlan(
        field_domains=field_domains,
        interned=tuple(planner.interned),
        series_expr=series_expr,
        series_key=series_key,
        scc_expr=scc_expr,
        scc_key=scc_key,
    )


def publish_attr_source(name: str, attr: str, value_expr: str) -> str:
    """Return a setattr publication line for generated function metadata."""
    return f"setattr({name}, {attr!r}, {value_expr})"


def key_domain_attr_source(
    name: str,
    *,
    keys: tuple[str, ...],
    domain_expr: str,
) -> str:
    """Return setattr lines that publish `{name}` `__key__` and `__domain__`."""
    return "\n".join(
        (
            publish_attr_source(name, "__key__", repr(keys)),
            publish_attr_source(name, "__domain__", domain_expr),
        )
    )


def constants_attr_source(name: str, constants: Sequence[str]) -> str:
    """Return a setattr line that publishes `{name}` `__constants__`."""
    return publish_attr_source(name, "__constants__", repr(tuple(constants)))
