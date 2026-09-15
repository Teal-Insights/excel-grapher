"""Named-axis primitives for inverted-tree codegen.

Lazy views, recurrence readers, and the `publish` metadata decorator. Excel
operators live in `excel.py`.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator, Mapping, Sequence
from itertools import product
from types import MappingProxyType
from typing import Any, Generic, Protocol, TypeVar, cast

from excel_grapher.core.grid import Range
from excel_grapher.core.types import FormulaValue
from excel_grapher.exporter.export_runtime.tensor import Axis, Domain, Tensor, TensorSchema
from excel_grapher.exporter.inverted_tree.excel import XlError, _as_number

T = TypeVar("T")
F = TypeVar("F", bound=Callable[..., object])
K = TypeVar("K", bound=tuple[object, ...])

TablePart = Callable[[], object] | Range


def lazy_table(rows: tuple[tuple[TablePart, ...], ...]) -> Range:
    """Expose a formula table without evaluating cells a lookup does not select.

    Each tuple is a horizontal strip of cell callbacks and views. Parts in a
    strip share height; a view contributes its shape in place, so a
    rectangular run of one series is one part instead of one callback per
    cell. Strips stack vertically and share the table width.
    """
    strips: list[tuple[int, int, list[tuple[int, TablePart]]]] = []
    width: int | None = None
    top = 1
    for strip in rows:
        parts: list[tuple[int, TablePart]] = []
        column = 0
        height: int | None = None
        for part in strip:
            if isinstance(part, Range):
                part_height, part_width = part.shape
                if height is None:
                    height = part_height
                elif height != part_height:
                    raise ValueError(
                        f"table strip parts differ in height: {height} and {part_height}"
                    )
                parts.append((column, part))
                column += part_width
            else:
                if height is None:
                    height = 1
                elif height != 1:
                    raise ValueError(f"table strip parts differ in height: {height} and 1")
                parts.append((column, part))
                column += 1
        height = height or 1
        if width is None:
            width = column
        elif width != column:
            raise ValueError(f"table rows differ in width: {width} and {column}")
        strips.append((top, height, parts))
        top += height

    def resolve(row: int, column: int) -> FormulaValue:
        for start_row, height, parts in strips:
            if start_row <= row < start_row + height:
                local_row = row - start_row + 1
                for start, part in reversed(parts):
                    if column - 1 >= start:
                        if isinstance(part, Range):
                            return part.cell(local_row, column - start)
                        return cast(FormulaValue, part())
                raise IndexError(column)
        raise IndexError(row)

    return Range("", 1, 1, top - 1 or 1, width or 1, lambda address: None, _coord_resolver=resolve)


class KeyedCompute(Protocol):
    """A generated `compute_*` or internals helper with published key metadata."""

    __key__: tuple[str, ...]
    __domain__: tuple[object, ...] | Domain | None
    __holes__: tuple[int, ...]


def publish(
    schema: TensorSchema | None = None,
    *,
    key: tuple[str, ...] | None = None,
    domain: object = None,
    holes: tuple[int, ...] = (),
    constants: Iterable[str] | None = None,
    cells: Mapping[K, str] | None = None,
) -> Callable[[F], F]:
    """Attach series metadata to a generated helper and return it unchanged.

    Sets `__key__`, `__domain__`, and `__holes__` on `fn`; a `schema` supplies
    the key fields and required domain of a tensor series. When `constants`
    is given, also sets `__constants__` to their names in sorted order.
    `cells` publishes immutable coordinate provenance as `__cells__`. Does
    not wrap `fn`.
    """
    if schema is not None:
        key = tuple(axis.name for axis in schema.domain.axes)
        domain = schema.domain
    if key is None:
        raise TypeError("publish needs a schema or explicit key fields")
    published_key = key

    def decorator(fn: F) -> F:
        target = cast(Any, fn)
        target.__key__ = published_key
        target.__domain__ = domain
        target.__holes__ = holes
        if cells is not None:
            target.__cells__ = MappingProxyType(dict(cells)) if isinstance(cells, dict) else cells
        if constants is not None:
            target.__constants__ = tuple(sorted(constants))
        return fn

    return decorator


def span(axis: Axis, first: object, last: object) -> tuple[str | int, ...]:
    """Return the keys of `axis` from `first` through `last`, inclusive.

    Lowers a worksheet range along one semantic axis. A key outside the axis
    raises `#REF!`.
    """
    try:
        start = axis.keys.index(cast(Any, first))
        stop = axis.keys.index(cast(Any, last))
    except ValueError as exc:
        raise XlError("#REF!") from exc
    if start > stop:
        raise XlError("#REF!")
    return axis.keys[start : stop + 1]


def view(
    values: Any,
    rows: Sequence[object] | Mapping[str, Sequence[object]] | None = None,
    cols: Sequence[object] | Mapping[str, Sequence[object]] | None = None,
    *,
    cols_first: bool = False,
) -> Range:
    """Expose a worksheet rectangle of one series without copying it.

    Each cell resolves `values[coordinate]` on access, so lookups evaluate
    only the cells they select and recurrence readers stay demand-driven.
    The coordinate is the row key followed by the column key; `cols_first`
    reverses that order, and an absent axis contributes no key.

    A block whose rows or columns nest several key fields selects keys per
    field name: `rows={"COUNTRY": keys, "SCENARIO": keys}` enumerates the
    product of those selections in worksheet order, and the coordinate is
    assembled in the order of the series' axes.
    """
    if isinstance(rows, Mapping) or isinstance(cols, Mapping):
        if not isinstance(rows or {}, Mapping) or not isinstance(cols or {}, Mapping):
            raise TypeError("view selects rows and columns by field name together")
        return _product_view(
            values,
            cast(Mapping[str, Sequence[object]], rows or {}),
            cast(Mapping[str, Sequence[object]], cols or {}),
        )
    row_keys: Sequence[object] = (None,) if rows is None else rows
    col_keys: Sequence[object] = (None,) if cols is None else cols

    def resolve(row: int, column: int) -> FormulaValue:
        parts = [row_keys[row - 1], col_keys[column - 1]]
        if cols_first:
            parts.reverse()
        coordinate = tuple(
            part
            for part, present in zip(
                parts, (rows, cols) if not cols_first else (cols, rows), strict=True
            )
            if present is not None
        )
        return cast(FormulaValue, values[coordinate])

    return Range(
        "",
        1,
        1,
        len(row_keys),
        len(col_keys),
        lambda address: None,
        _coord_resolver=resolve,
    )


def _product_view(
    values: Any,
    rows: Mapping[str, Sequence[object]],
    cols: Mapping[str, Sequence[object]],
) -> Range:
    """Lazy block over the product of per-field key selections."""
    names = tuple(axis.name for axis in values.domain.axes)
    if set(rows) | set(cols) != set(names) or set(rows) & set(cols):
        raise ValueError(f"view selections must cover the axes {names!r} once each")
    fields = (*rows, *cols)
    row_products: list[tuple[object, ...]] = list(product(*[tuple(keys) for keys in rows.values()]))
    col_products: list[tuple[object, ...]] = list(product(*[tuple(keys) for keys in cols.values()]))

    def resolve(row: int, column: int) -> FormulaValue:
        selected = (*row_products[row - 1], *col_products[column - 1])
        keys = dict(zip(fields, selected, strict=True))
        return cast(FormulaValue, values[tuple(keys[name] for name in names)])

    return Range(
        "",
        1,
        1,
        len(row_products),
        len(col_products),
        lambda address: None,
        _coord_resolver=resolve,
    )


def at_anchor(value: T, rows: object, cols: object) -> T:
    """Return a scalar `OFFSET` anchor when both displacements are zero.

    A scalar series has no other bound cells to move to, so a non-zero
    displacement raises `#VALUE!` like a positional read past the series.
    """
    if int(_as_number(rows)) != 0 or int(_as_number(cols)) != 0:
        raise XlError("#VALUE!")
    return value


def axis_step(axis: Axis, key: object, steps: object) -> str | int:
    """Return the key `steps` positions after `key` along `axis`.

    Lowers `OFFSET` moves along one worksheet axis. A position outside the
    bound series raises `#VALUE!`, matching positional `xl_at` selection.
    """
    try:
        position = axis.keys.index(cast(Any, key)) + int(_as_number(steps))
    except ValueError as exc:
        raise XlError("#VALUE!") from exc
    if position < 0 or position >= len(axis.keys):
        raise XlError("#VALUE!")
    return axis.keys[position]


def require_aligned(*series: Sequence[object]) -> int:
    """Return the common length, or fail if any series has a different length."""
    if not series:
        raise ValueError("require_aligned expected at least one series")
    lengths = [len(item) for item in series]
    if len(set(lengths)) != 1:
        raise ValueError(f"misaligned series lengths: {lengths}")
    return lengths[0]


def take(values: Sequence[T], indices: Sequence[int] | slice) -> tuple[T, ...]:
    """Return `values` at 0-based `indices`, failing closed on out-of-range.

    `indices` may be a sequence (including `range`) or a `slice`. A slice
    expands with `range(start, stop, step)`: omitted `start` is 0, omitted
    `stop` is `len(values)`, and `stop` is not clamped — an explicit stop
    past the series length fails closed, same as a tuple of those indices.

    The orchestrator gathers from a catalog-order array into a dense working
    buffer. Internals zip/scan that buffer; they never see holes.
    """
    if isinstance(indices, slice):
        start = 0 if indices.start is None else indices.start
        stop = len(values) if indices.stop is None else indices.stop
        step = 1 if indices.step is None else indices.step
        if step == 0:
            raise ValueError("take slice step cannot be zero")
        indices = range(start, stop, step)
    length = len(values)
    result: list[T] = []
    for index in indices:
        if index < 0 or index >= length:
            raise ValueError(f"take index {index} is outside series of length {length}")
        result.append(values[index])
    return tuple(result)


def as_records(
    compute: KeyedCompute,
    result: object,
    *,
    measure: str = "OBS_VALUE",
) -> list[dict[str, object]]:
    """Zip a compute result with `__key__` / `__domain__` into records.

    Named results iterate by coordinate identity. Legacy metadata requires an
    explicitly ordered sequence and remains supported by this record adapter.
    """
    keys = compute.__key__
    domain = compute.__domain__
    if isinstance(domain, Domain):
        if not isinstance(result, Tensor) or result.domain != domain:
            raise ValueError("result must be a Tensor over the declared result domain")
        return [
            dict(zip(keys, coord, strict=True)) | {measure: value}
            for coord, value in result.items()
        ]
    if domain is None:
        return [{measure: result}]
    if not isinstance(result, Sequence):
        raise ValueError("legacy result metadata requires an explicitly ordered sequence")
    if len(domain) != len(result):
        raise ValueError(f"result length {len(result)} does not match domain length {len(domain)}")
    records: list[dict[str, object]] = []
    for point, value in zip(domain, result, strict=True):
        if not keys:
            record: dict[str, object] = {measure: value}
        elif len(keys) == 1:
            record = {keys[0]: point, measure: value}
        else:
            if not isinstance(point, tuple) or len(point) != len(keys):
                raise ValueError(f"domain point {point!r} does not match key {keys!r}")
            record = dict(zip(keys, point, strict=True))
            record[measure] = value
        records.append(record)
    return records


class InstanceCycleError(ValueError):
    """Demand-driven evaluation hit a same-index circular reference."""


class CoordinateReader(Generic[T]):
    """Private demand-driven series; only completed tensors are published.

    A read of a valid coordinate that is absent from a sparse required
    domain is a blank (`None`), matching a structural hole Excel evaluates
    as empty rather than as an out-of-domain error.
    """

    def __init__(
        self,
        series_id: str,
        domain: Domain,
        compute: Callable[..., T],
    ) -> None:
        self._series_id = series_id
        self._domain = domain
        self._compute = compute
        self._memo: dict[tuple[str, tuple[str | int, ...]], T] = {}
        self._active: set[tuple[str, tuple[str | int, ...]]] = set()

    @property
    def domain(self) -> Domain:
        """Coordinates this reader can compute."""
        return self._domain

    def __getitem__(self, key: str | int | tuple[str | int, ...]) -> T:
        coordinate = key if isinstance(key, tuple) else (key,)
        if self._domain.sparse_get(coordinate) is None:
            return cast(T, None)
        identity = (self._series_id, coordinate)
        if identity in self._memo:
            return self._memo[identity]
        if identity in self._active:
            raise InstanceCycleError(f"circular reference at {self._series_id}{coordinate!r}")
        self._active.add(identity)
        try:
            try:
                value = self._compute(*coordinate)
            except XlError as error:
                value = cast(T, error.code)
            self._memo[identity] = value
            return value
        finally:
            self._active.remove(identity)


def evaluate(
    formula: Callable[..., T], domain: Iterable[tuple[str | int, ...]]
) -> Iterator[tuple[tuple[str | int, ...], T | str]]:
    """Apply `formula` to every coordinate of `domain`, storing Excel errors as codes."""
    for coordinate in domain:
        try:
            yield coordinate, formula(*coordinate)
        except XlError as error:
            yield coordinate, error.code


def eval_instance(
    statement: str,
    index: int,
    compute: Callable[[int], T],
    memo: dict[tuple[str, int], T],
    stack: set[tuple[str, int]],
) -> T:
    """Return the memoized value of `statement` at catalog `index`.

    This is the rung-3 dispatcher: demand-driven instance evaluation with an
    on-stack set that raises `InstanceCycleError` on a real cycle.
    """
    if index < 0:
        raise XlError("#REF!")
    key = (statement, index)
    if key in memo:
        return memo[key]
    if key in stack:
        raise InstanceCycleError(f"distance-zero cycle evaluating {statement}[{index}]")
    stack.add(key)
    try:
        value = compute(index)
    finally:
        stack.remove(key)
    memo[key] = value
    return value
