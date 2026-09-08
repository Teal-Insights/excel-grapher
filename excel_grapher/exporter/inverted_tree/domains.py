"""Key-domain tuples for inverted-tree `data.py` and compute metadata.

Field domains (`TIME_PERIOD_DOMAIN`, ...) are the catalog-order union of resolved
key values, one tuple per distinct field. Each `compute_*` / internals helper
publishes `__key__` and `__domain__` so callers index by key instead of column
count (#676). Generated modules attach that metadata with `@publish` (#766).
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from dataclasses import field as data_field
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


_MAX_PRODUCT_SCAN = 250_000
_MAX_EXCEPTIONS = 8


def _product_comprehension(
    fields: tuple[str, ...],
    field_exprs: Sequence[str],
    *,
    exclude: Sequence[str] = (),
) -> str:
    used: set[str] = set()
    names = [_loop_var(field, used) for field in fields]
    tuple_body = ", ".join(names)
    gens = " ".join(f"for {name} in {expr}" for name, expr in zip(names, field_exprs, strict=True))
    if not exclude:
        return f"tuple(({tuple_body}) {gens})"
    if len(exclude) == 1:
        return f"tuple(({tuple_body}) {gens} if ({tuple_body}) != {exclude[0]})"
    inner = ", ".join(exclude)
    return f"tuple(({tuple_body}) {gens} if ({tuple_body}) not in {{{inner}}})"


def _maybe_product(axes: Sequence[tuple[object, ...]]) -> tuple[object, ...] | None:
    """Return `tuple(product(*axes))` when the product is small enough to scan."""
    size = 1
    for axis in axes:
        if not axis:
            return ()
        size *= len(axis)
        if size > _MAX_PRODUCT_SCAN:
            return None
    return tuple(product(*axes))


def _worth_emitting(source: str, points: tuple[object, ...]) -> bool:
    """True when `source` is shorter than a tuple literal of `points`."""
    if not points:
        return False
    sample_n = min(8, len(points))
    sample = sum(len(repr(point)) + 2 for point in points[:sample_n]) / sample_n
    estimated = int(sample * len(points)) + 2
    return len(source) < estimated


def _cover_rectangles(
    points: tuple[tuple[object, ...], ...],
) -> list[tuple[tuple[object, ...], ...]] | None:
    """Split `points` into consecutive Cartesian blocks, preserving order.

    Each block is a tuple of per-dimension value tuples such that
    `product(*block)` reproduces that block's points. Returns `None` when
    `points` are not uniform tuples.
    """
    if not points:
        return []
    ndim = len(points[0])
    if any(not isinstance(point, tuple) or len(point) != ndim for point in points):
        return None
    if ndim == 1:
        return [(tuple(point[0] for point in points),)]

    groups: list[tuple[object, tuple[tuple[object, ...], ...]]] = []
    index = 0
    while index < len(points):
        outer = points[index][0]
        inner: list[tuple[object, ...]] = []
        while index < len(points) and points[index][0] == outer:
            inner.append(points[index][1:])
            index += 1
        groups.append((outer, tuple(inner)))

    covered: list[tuple[object, list[tuple[tuple[object, ...], ...]]]] = []
    for outer, inner_points in groups:
        inner_blocks = _cover_rectangles(inner_points)
        if inner_blocks is None:
            return None
        covered.append((outer, inner_blocks))

    blocks: list[tuple[tuple[object, ...], ...]] = []
    group_index = 0
    while group_index < len(covered):
        outer, inner_blocks = covered[group_index]
        if len(inner_blocks) == 1:
            inner_axes = inner_blocks[0]
            outers = [outer]
            group_index += 1
            while (
                group_index < len(covered)
                and len(covered[group_index][1]) == 1
                and covered[group_index][1][0] == inner_axes
            ):
                outers.append(covered[group_index][0])
                group_index += 1
            blocks.append((tuple(outers), *inner_axes))
        else:
            for inner_axes in inner_blocks:
                blocks.append(((outer,), *inner_axes))
            group_index += 1

    rebuilt: list[object] = []
    for block in blocks:
        rebuilt.extend(product(*block))
    if tuple(rebuilt) != points:
        return None
    return blocks


def _triangle_source(
    keys: tuple[str, ...],
    points: tuple[object, ...],
    per_field: Sequence[tuple[object, ...]],
    axis_exprs: Sequence[str],
) -> str | None:
    """Return an enumerate comprehension for a 2-D triangular domain."""
    if len(keys) != 2 or len(per_field) != 2:
        return None
    groups: list[tuple[object, tuple[object, ...]]] = []
    index = 0
    while index < len(points):
        point = points[index]
        if not isinstance(point, tuple) or len(point) != 2:
            return None
        outer = point[0]
        inners: list[object] = []
        while index < len(points):
            current = points[index]
            if not isinstance(current, tuple) or current[0] != outer:
                break
            inners.append(current[1])
            index += 1
        groups.append((outer, tuple(inners)))
    outers = tuple(group[0] for group in groups)
    if outers != per_field[0]:
        return None
    inner_full = per_field[1]
    count = len(groups)
    got = tuple(group[1] for group in groups)
    used: set[str] = set()
    outer_var = _loop_var(keys[0], used)
    inner_var = _loop_var(keys[1], used)
    idx = "i"
    while idx in used:
        idx = f"{idx}x"
    outer_expr, inner_expr = axis_exprs
    if got == tuple(inner_full[: i + 1] for i in range(count)):
        return (
            f"tuple(({outer_var}, {inner_var}) for {idx}, {outer_var} in enumerate({outer_expr}) "
            f"for {inner_var} in {inner_expr}[: {idx} + 1])"
        )
    if got == tuple(inner_full[: count - i] for i in range(count)):
        return (
            f"tuple(({outer_var}, {inner_var}) for {idx}, {outer_var} in enumerate({outer_expr}) "
            f"for {inner_var} in {inner_expr}[: {count} - {idx}])"
        )
    if got == tuple(inner_full[i:] for i in range(count)):
        return (
            f"tuple(({outer_var}, {inner_var}) for {idx}, {outer_var} in enumerate({outer_expr}) "
            f"for {inner_var} in {inner_expr}[{idx}:])"
        )
    if got == tuple(inner_full[count - i - 1 :] for i in range(count)):
        return (
            f"tuple(({outer_var}, {inner_var}) for {idx}, {outer_var} in enumerate({outer_expr}) "
            f"for {inner_var} in {inner_expr}[{count} - {idx} - 1 :])"
        )
    return None


def _product_variant_source(
    keys: tuple[str, ...],
    points: tuple[object, ...],
    per_field: Sequence[tuple[object, ...]],
    axis_exprs: Sequence[str],
    point_expr: Callable[[tuple[str, ...], object], str],
) -> str | None:
    """Return a product comprehension, slice, or small-exception filter."""
    generated = _maybe_product(per_field)
    if generated is None:
        return None
    comprehension = _product_comprehension(keys, axis_exprs)
    if generated == points:
        return comprehension
    extra = len(generated) - len(points)
    if extra <= 0 or extra > _MAX_EXCEPTIONS:
        return None
    if generated[: len(points)] == points:
        return f"{comprehension}[:-1]" if extra == 1 else f"{comprehension}[:-{extra}]"
    if generated[-len(points) :] == points:
        return f"{comprehension}[{extra}:]"
    missing: list[object] = []
    gen_iter = iter(generated)
    try:
        current = next(gen_iter)
        for point in points:
            while current != point:
                missing.append(current)
                current = next(gen_iter)
            current = next(gen_iter, _MISSING)
        if current is not _MISSING:
            missing.append(current)
            missing.extend(gen_iter)
    except StopIteration:
        return None
    if len(missing) != extra:
        return None
    exclude = [point_expr(keys, item) for item in missing]
    return _product_comprehension(keys, axis_exprs, exclude=exclude)


_MISSING = object()


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


class _Planner:
    """Build domain expressions, interning tuples that are not slices/products."""

    def __init__(self, field_domains: dict[str, tuple[Scalar, ...]]) -> None:
        self.field_domains = field_domains
        self.interned: list[tuple[str, tuple[object, ...]]] = []
        self.interned_source: dict[str, str] = {}

    def intern(self, points: tuple[object, ...], source: str | None = None) -> str:
        for name, values in self.interned:
            if values == points:
                return f"data.{name}"
        name = f"_DOMAIN_{len(self.interned)}"
        self.interned.append((name, points))
        if source is not None:
            self.interned_source[name] = source
        return f"data.{name}"

    def field_ref(
        self, field: str, values: tuple[object, ...], *, qualified: bool = True
    ) -> str | None:
        full = self.field_domains.get(field)
        if full is None:
            return None
        slc = _contiguous_slice(full, values)
        if slc is None:
            return None
        name = domain_const_name(field)
        ref = f"data.{name}" if qualified else name
        if slc == slice(None):
            return ref
        return f"{ref}{_slice_source(slc)}"

    def axis_expr(self, field: str, values: tuple[object, ...], *, qualified: bool) -> str:
        ref = self.field_ref(field, values, qualified=qualified)
        return ref if ref is not None else repr(values)

    def value_expr(self, field: str, value: object, *, qualified: bool) -> str:
        full = self.field_domains.get(field)
        if full is None:
            return repr(value)
        try:
            index = full.index(value)
        except ValueError:
            return repr(value)
        name = domain_const_name(field)
        ref = f"data.{name}" if qualified else name
        if index == len(full) - 1:
            return f"{ref}[-1]"
        return f"{ref}[{index}]"

    def point_expr(self, keys: tuple[str, ...], point: object, *, qualified: bool) -> str:
        if not isinstance(point, tuple) or len(point) != len(keys):
            return repr(point)
        parts = [
            self.value_expr(field, value, qualified=qualified)
            for field, value in zip(keys, point, strict=True)
        ]
        return f"({', '.join(parts)})"

    def _block_source(
        self,
        keys: tuple[str, ...],
        block: tuple[tuple[object, ...], ...],
        *,
        qualified: bool,
    ) -> tuple[str, int]:
        count = 1
        for axis in block:
            count *= len(axis)
        if count == 1:
            point = tuple(axis[0] for axis in block)
            return self.point_expr(keys, point, qualified=qualified), 1
        exprs = [
            self.axis_expr(field, axis, qualified=qualified)
            for field, axis in zip(keys, block, strict=True)
        ]
        return _product_comprehension(keys, exprs), count

    def _blocks_source(
        self,
        keys: tuple[str, ...],
        blocks: list[tuple[tuple[object, ...], ...]],
        *,
        qualified: bool,
    ) -> str | None:
        parts: list[tuple[str, int]] = []
        for block in blocks:
            if len(block) != len(keys):
                return None
            parts.append(self._block_source(keys, block, qualified=qualified))
        if len(parts) == 1:
            return parts[0][0]
        bits = [expr if n == 1 else f"*{expr}" for expr, n in parts]
        return f"({', '.join(bits)})"

    def _compact_source(
        self,
        keys: tuple[str, ...],
        points: tuple[object, ...],
        per_field: list[tuple[object, ...]],
        *,
        qualified: bool,
    ) -> str | None:
        axis_exprs = [
            self.axis_expr(field, values, qualified=qualified)
            for field, values in zip(keys, per_field, strict=True)
        ]
        candidates: list[str] = []
        product_form = _product_variant_source(
            keys,
            points,
            per_field,
            axis_exprs,
            lambda ks, pt: self.point_expr(ks, pt, qualified=qualified),
        )
        if product_form is not None:
            candidates.append(product_form)
        triangle = _triangle_source(keys, points, per_field, axis_exprs)
        if triangle is not None:
            candidates.append(triangle)
        tuple_points = tuple(point for point in points if isinstance(point, tuple))
        if len(tuple_points) == len(points):
            blocks = _cover_rectangles(tuple_points)
            if blocks is not None:
                block_expr = self._blocks_source(keys, blocks, qualified=qualified)
                if block_expr is not None:
                    candidates.append(block_expr)
        if not candidates:
            return None
        return min(candidates, key=len)

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
        generated = _maybe_product(per_field)
        if generated == points:
            refs: list[str] = []
            for field, values in zip(keys, per_field, strict=True):
                ref = self.field_ref(field, values)
                if ref is None:
                    break
                refs.append(ref)
            else:
                return _product_comprehension(keys, refs)
        source = self._compact_source(keys, points, per_field, qualified=False)
        if source is not None and _worth_emitting(source, points):
            return self.intern(points, source=source)
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
        interned_source=dict(planner.interned_source),
    )


def publish_decorator_source(
    *,
    keys: tuple[str, ...],
    domain_expr: str,
    holes: tuple[int, ...] = (),
    constants: Sequence[str] | None = None,
) -> str:
    """Return a `@publish(...)` decorator for generated series metadata.

    Empty `holes` are omitted (`holes=()` is the decorator default). `constants`
    is emitted only when given, including an empty tuple for api `compute_*`.
    """
    args = [f"key={keys!r}", f"domain={domain_expr}"]
    if holes:
        args.append(f"holes={holes!r}")
    if constants is not None:
        args.append(f"constants={tuple(constants)!r}")
    body = ",\n    ".join(args)
    return f"@publish(\n    {body},\n)"
