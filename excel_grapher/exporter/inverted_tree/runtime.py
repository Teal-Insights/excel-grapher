"""Excel operators and series-alignment primitives for inverted-tree codegen.

Mechanical extraction emits calls to these helpers instead of reading cells
from an evaluation context. `take` gathers catalog-order series by index;
internals never see holes and never fetch extra items to pad a result.

A **measure** is a numeric observation or an Excel error code string
(`#REF!`, `#DIV/0!`, …) — the same `err.code` ctx stores on Records. Operators
raise `XlError`; series-member loops catch it so one error cell does not abort
the tuple.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from datetime import date, datetime
from itertools import product
from types import MappingProxyType
from typing import Any, Generic, Literal, NoReturn, Protocol, TypeGuard, TypeVar, cast, overload

from excel_grapher.core.grid import Range
from excel_grapher.core.logic_funcs import logical_and, logical_if, logical_not, logical_or
from excel_grapher.core.lookup_funcs import index_cells, match_cells, vlookup_cells
from excel_grapher.core.math_funcs import average_cells, exp_number, max_cells, min_cells, sum_cells
from excel_grapher.core.operators import xl_add as _core_add
from excel_grapher.core.operators import xl_concat as _core_concat
from excel_grapher.core.operators import xl_div as _core_div
from excel_grapher.core.operators import xl_eq as _core_eq
from excel_grapher.core.operators import xl_ge as _core_ge
from excel_grapher.core.operators import xl_gt as _core_gt
from excel_grapher.core.operators import xl_le as _core_le
from excel_grapher.core.operators import xl_lt as _core_lt
from excel_grapher.core.operators import xl_mul as _core_mul
from excel_grapher.core.operators import xl_ne as _core_ne
from excel_grapher.core.operators import xl_neg as _core_neg
from excel_grapher.core.operators import xl_pos as _core_pos
from excel_grapher.core.operators import xl_pow as _core_pow
from excel_grapher.core.operators import xl_sub as _core_sub
from excel_grapher.core.sumproduct import sumproduct_cells
from excel_grapher.core.types import CellValue, FormulaValue
from excel_grapher.core.types import XlError as CoreXlError
from excel_grapher.core.types import XlErrorException as SharedXlError
from excel_grapher.exporter.export_runtime.error_funcs import xl_iferror as _shared_iferror
from excel_grapher.exporter.export_runtime.error_funcs import xl_ifna as _shared_ifna
from excel_grapher.exporter.export_runtime.error_funcs import xl_isblank as _shared_isblank
from excel_grapher.exporter.export_runtime.error_funcs import xl_iserror as _shared_iserror
from excel_grapher.exporter.export_runtime.error_funcs import xl_isna as _shared_isna
from excel_grapher.exporter.export_runtime.error_funcs import xl_isnumber as _shared_isnumber
from excel_grapher.exporter.export_runtime.lookup import xl_hlookup as _shared_hlookup
from excel_grapher.exporter.export_runtime.lookup import xl_lookup as _shared_lookup
from excel_grapher.exporter.export_runtime.lookup import xl_xlookup as _shared_xlookup
from excel_grapher.exporter.export_runtime.math import xl_countif as _shared_countif
from excel_grapher.exporter.export_runtime.math import xl_large as _shared_large
from excel_grapher.exporter.export_runtime.math import xl_npv as _shared_npv
from excel_grapher.exporter.export_runtime.math import xl_rank as _shared_rank
from excel_grapher.exporter.export_runtime.math import xl_round as _shared_round
from excel_grapher.exporter.export_runtime.math import xl_rounddown as _shared_rounddown
from excel_grapher.exporter.export_runtime.math import xl_stdev as _shared_stdev
from excel_grapher.exporter.export_runtime.tensor import Axis, Domain, Tensor, TensorSchema
from excel_grapher.exporter.export_runtime.text import xl_left as _shared_left
from excel_grapher.exporter.export_runtime.text import xl_numbervalue as _shared_numbervalue
from excel_grapher.series_bindings.input_coerce import (
    apply_input_value_map as apply_input_value_map,
)
from excel_grapher.series_bindings.input_coerce import require_input_domain as require_input_domain

T = TypeVar("T")
F = TypeVar("F", bound=Callable[..., object])
K = TypeVar("K", bound=tuple[object, ...])


TablePart = Callable[[], object] | Range


def lazy_table(rows: tuple[tuple[TablePart, ...], ...]) -> Range:
    """Expose a formula table without evaluating cells a lookup does not select.

    Each row lists cell callbacks and one-row views side by side; a view
    contributes its columns in place, so a run of one series' cells is one
    part instead of one callback per cell.
    """
    layout: list[list[tuple[int, TablePart]]] = []
    width: int | None = None
    for row in rows:
        parts: list[tuple[int, TablePart]] = []
        column = 0
        for part in row:
            parts.append((column, part))
            if isinstance(part, Range):
                if part.shape[0] != 1:
                    raise ValueError("table rows accept one-row views")
                column += part.shape[1]
            else:
                column += 1
        if width is None:
            width = column
        elif width != column:
            raise ValueError(f"table rows differ in width: {width} and {column}")
        layout.append(parts)

    def resolve(row: int, column: int) -> FormulaValue:
        for start, part in reversed(layout[row - 1]):
            if column - 1 >= start:
                if isinstance(part, Range):
                    return part.cell(1, column - start)
                return cast(FormulaValue, part())
        raise IndexError(column)

    return Range("", 1, 1, len(rows), width or 1, lambda address: None, _coord_resolver=resolve)


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
    constants: tuple[str, ...] | None = None,
    cells: Mapping[K, str] | None = None,
) -> Callable[[F], F]:
    """Attach series metadata to a generated helper and return it unchanged.

    Sets `__key__`, `__domain__`, and `__holes__` on `fn`; a `schema` supplies
    the key fields and required domain of a tensor series. When `constants`
    is given, also sets `__constants__`. `cells` publishes immutable
    coordinate provenance as `__cells__`. Does not wrap `fn`.
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
            target.__constants__ = constants
        return fn

    return decorator


XL_ERROR_CODES = frozenset(
    {
        "#VALUE!",
        "#REF!",
        "#DIV/0!",
        "#N/A",
        "#NAME?",
        "#NUM!",
        "#NULL!",
    }
)


class XlError(Exception):
    """Excel error value raised as a Python exception."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


def is_error(value: object) -> TypeGuard[str]:
    """True when `value` is an Excel error code string."""
    return isinstance(value, str) and value in XL_ERROR_CODES


@overload
def as_measure(value: object, dtype: Literal["float"] = "float") -> float | str: ...
@overload
def as_measure(value: object, dtype: Literal["int"]) -> int | str: ...
@overload
def as_measure(value: object, dtype: Literal["str"]) -> str: ...
@overload
def as_measure(value: object, dtype: Literal["bool"]) -> bool | str: ...
@overload
def as_measure(value: object, dtype: Literal["datetime"]) -> datetime | str: ...
def as_measure(value: object, dtype: str = "float") -> int | float | str | bool | datetime | None:
    """Coerce a helper result to a measure: number or cached text.

    Operators still raise `XlError`. Series-member boundaries catch that and
    store `err.code` here so a `#REF!` cell does not abort the rest of a series.
    Non-numeric cached strings (`n/a`, `..`) pass through as measures.
    Blank cells (`None`) stay `None`.

    Overloads narrow the return by `dtype`: the default `float` path is
    `float | str` so generated `list[float | str]` accumulators type-check.
    """
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, XlError):
        return value.code
    if dtype == "int":
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            return int(value)
        raise TypeError(f"cannot coerce {type(value).__name__} to int measure")
    if dtype == "str":
        return str(value)
    if dtype == "bool":
        return bool(value)
    if dtype == "datetime":
        if isinstance(value, datetime):
            return value
        if isinstance(value, date):
            return datetime(value.year, value.month, value.day)
        raise TypeError(f"cannot coerce {type(value).__name__} to datetime measure")
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, int | float):
        return float(value)
    raise TypeError(f"cannot coerce {type(value).__name__} to float measure")


def _raise_stored_error(value: object) -> None:
    """Re-raise a cached Excel error-code measure."""
    if isinstance(value, str) and is_error(value):
        raise XlError(value)


def _raise_stored_errors_in(value: object) -> None:
    """Re-raise stored error-code measures in scalars, sequences, and views."""
    if isinstance(value, Range):
        for item in value.iter_values():
            _raise_stored_error(item)
        return
    if isinstance(value, str) or not isinstance(value, Sequence):
        _raise_stored_error(value)
        return
    for item in value:
        _raise_stored_errors_in(item)


def _as_core_cells(value: object) -> CellValue:
    """Convert stored error-code measures to core sentinels for shared helpers."""
    if isinstance(value, str):
        converted = CoreXlError.from_text(value)
        return converted if converted is not None else value
    if isinstance(value, Sequence):
        return cast(CellValue, [_as_core_cells(item) for item in value])
    return cast(CellValue, value)


def _adapt_core(value: object) -> object:
    """Raise `XlError` when `core` returned a sentinel."""
    if isinstance(value, CoreXlError):
        raise XlError(value.value)
    return value


def _as_formula(value: object) -> FormulaValue:
    """Narrow a generated-code operand to a `core` formula value."""
    return cast(FormulaValue, value)


def _arith_operand(value: object) -> FormulaValue:
    """Prepare an arithmetic operand for `core` (blank text is `0`)."""
    _raise_stored_error(value)
    if isinstance(value, str) and value.replace("\u00a0", "").strip() == "":
        return 0.0
    return _as_formula(value)


def _as_number(value: object) -> float:
    """Coerce `value` via core `to_number`, re-raising stored error codes."""
    from excel_grapher.core.coercions import to_number

    number = to_number(_arith_operand(value))
    if isinstance(number, CoreXlError):
        raise XlError(number.value)
    return float(number)


def xl_add(left: object, right: object) -> object:
    """Excel `+` via `core.operators.xl_add`."""
    return _adapt_core(_core_add(_arith_operand(left), _arith_operand(right)))


def xl_concat(left: object, right: object) -> object:
    """Excel `&` with shared blank, boolean, number, and error semantics."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_concat(_as_formula(left), _as_formula(right)))


def xl_sub(left: object, right: object) -> object:
    """Excel `-` via `core.operators.xl_sub`."""
    return _adapt_core(_core_sub(_arith_operand(left), _arith_operand(right)))


def xl_mul(left: object, right: object) -> object:
    """Excel `*` via `core.operators.xl_mul`."""
    return _adapt_core(_core_mul(_arith_operand(left), _arith_operand(right)))


def xl_div(numerator: object, denominator: object) -> object:
    """Excel `/` via `core.operators.xl_div`."""
    return _adapt_core(_core_div(_arith_operand(numerator), _arith_operand(denominator)))


def xl_pow(left: object, right: object) -> object:
    """Excel `^` via `core.operators.xl_pow`."""
    return _adapt_core(_core_pow(_arith_operand(left), _arith_operand(right)))


def xl_neg(value: object) -> object:
    """Excel unary `-` via `core.operators.xl_neg`."""
    return _adapt_core(_core_neg(_arith_operand(value)))


def xl_pos(value: object) -> object:
    """Excel unary `+` via `core.operators.xl_pos`."""
    return _adapt_core(_core_pos(_arith_operand(value)))


def xl_eq(left: object, right: object) -> object:
    """Excel `=` via `core.operators.xl_eq`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_eq(_as_formula(left), _as_formula(right)))


def xl_ne(left: object, right: object) -> object:
    """Excel `<>` via `core.operators.xl_ne`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_ne(_as_formula(left), _as_formula(right)))


def xl_lt(left: object, right: object) -> object:
    """Excel `<` via `core.operators.xl_lt`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_lt(_as_formula(left), _as_formula(right)))


def xl_gt(left: object, right: object) -> object:
    """Excel `>` via `core.operators.xl_gt`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_gt(_as_formula(left), _as_formula(right)))


def xl_le(left: object, right: object) -> object:
    """Excel `<=` via `core.operators.xl_le`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_le(_as_formula(left), _as_formula(right)))


def xl_ge(left: object, right: object) -> object:
    """Excel `>=` via `core.operators.xl_ge`."""
    _raise_stored_error(left)
    _raise_stored_error(right)
    return _adapt_core(_core_ge(_as_formula(left), _as_formula(right)))


OPERATOR_TABLE = {
    "+": xl_add,
    "-": xl_sub,
    "*": xl_mul,
    "/": xl_div,
    "^": xl_pow,
    "=": xl_eq,
    "<>": xl_ne,
    "<": xl_lt,
    ">": xl_gt,
    "<=": xl_le,
    ">=": xl_ge,
    "-u": xl_neg,
    "+u": xl_pos,
}


def xl_exp(*args: object) -> object:
    """Excel `EXP` via `core.math_funcs.exp_number`."""
    for arg in args:
        _raise_stored_error(arg)
    return _adapt_core(exp_number(*cast(tuple[CellValue, ...], args)))


def xl_sum(*args: object) -> object:
    """Excel `SUM` via `core.math_funcs.sum_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(sum_cells(*(_as_core_cells(arg) for arg in args)))


def xl_average(*args: object) -> object:
    """Excel `AVERAGE` via `core.math_funcs.average_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(average_cells(*(_as_core_cells(arg) for arg in args)))


def xl_min(*args: object) -> object:
    """Excel `MIN` via `core.math_funcs.min_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(min_cells(*(_as_core_cells(arg) for arg in args)))


def xl_max(*args: object) -> object:
    """Excel `MAX` via `core.math_funcs.max_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(max_cells(*(_as_core_cells(arg) for arg in args)))


def xl_if(cond: object, then_value: object, else_value: object = False) -> object:
    """Excel `IF` via `core.logic_funcs.logical_if` (scalar or element-wise)."""
    return _adapt_core(logical_if(cond, then_value, else_value))


def xl_and(*args: object) -> object:
    """Excel `AND` via `core.logic_funcs.logical_and`."""
    return _adapt_core(logical_and(*(_as_core_cells(arg) for arg in args)))


def xl_or(*args: object) -> object:
    """Excel `OR` via `core.logic_funcs.logical_or`."""
    return _adapt_core(logical_or(*(_as_core_cells(arg) for arg in args)))


def xl_not(arg: object) -> object:
    """Excel `NOT` via `core.logic_funcs.logical_not`."""
    return _adapt_core(logical_not(_as_core_cells(arg)))


def xl_sumproduct(*args: object) -> object:
    """Excel `SUMPRODUCT` via `core.sumproduct.sumproduct_cells`."""
    for arg in args:
        _raise_stored_errors_in(arg)
    return _adapt_core(sumproduct_cells(*(_as_core_cells(arg) for arg in args)))


def xl_choose(index: object, *choices: float) -> float:
    """Excel `CHOOSE`: 1-based selection over already-evaluated arguments."""
    position = int(_as_number(index))
    if position < 1 or position > len(choices):
        raise XlError("#VALUE!")
    return choices[position - 1]


def xl_choose_lazy(index: object, *choices: Callable[[], object]) -> object:
    """Select one CHOOSE branch before evaluating its workbook expression."""
    selected = int(xl_choose(index, *range(len(choices))))
    return choices[selected]()


def xl_choose_range(index: object, cells: Range) -> object:
    """Select the `index`-th cell of a one-row or one-column view, as `CHOOSE` lists it."""
    rows, cols = cells.shape
    selected = int(xl_choose(index, *range(rows * cols)))
    return cells.cell(selected // cols + 1, selected % cols + 1)


def xl_index(array: object, row_num: object = None, col_num: object = None) -> object:
    """Excel `INDEX` via `core.lookup_funcs.index_cells`."""
    return _adapt_core(index_cells(array, row_num, col_num))


def xl_match(lookup: object, lookup_array: object, match_type: int = 0) -> int:
    """Excel `MATCH` via `core.lookup_funcs.match_cells`."""
    _raise_stored_error(lookup)
    result = match_cells(lookup, _as_core_cells(lookup_array), match_type)
    adapted = _adapt_core(result)
    if not isinstance(adapted, int | float):
        raise TypeError(f"MATCH returned {type(adapted).__name__}")
    return int(adapted)


def xl_vlookup(
    lookup: object,
    table_array: object,
    col_index_num: object,
    range_lookup: object = True,
) -> object:
    """Excel `VLOOKUP` via `core.lookup_funcs.vlookup_cells`."""
    _raise_stored_error(lookup)
    _raise_stored_error(col_index_num)
    _raise_stored_error(range_lookup)
    return _adapt_core(vlookup_cells(lookup, table_array, col_index_num, range_lookup))


def xl_isnumber(value: object) -> bool:
    """Excel `ISNUMBER`: True only for non-bool numbers; False for blanks and errors."""
    if isinstance(value, str) and is_error(value):
        return False
    return not isinstance(value, bool) and isinstance(value, int | float)


def _call_shared(function: Callable[..., object], *args: object) -> object:
    """Translate the shared runtime's exception channel at the export boundary."""
    try:
        return function(*args)
    except SharedXlError as exc:
        raise XlError(exc.code.value) from exc


def _shared_value(function: Callable[..., object], *args: object) -> object:
    """Pass scalar and range values to a shared worksheet function."""
    return _call_shared(function, *(_shared_operand(arg) for arg in args))


def _shared_operand(value: object) -> object:
    """Preserve error cells for each shared function to consume or propagate."""
    if isinstance(value, str):
        return CoreXlError(value) if is_error(value) else value
    if isinstance(value, Sequence):
        return tuple(_shared_operand(item) for item in value)
    return value


def _shared_thunk(value: Callable[[], object]) -> Callable[[], object]:
    """Expose stored or raised inverted-tree errors to a shared lazy consumer."""

    def evaluate() -> object:
        try:
            result = value()
            _raise_stored_error(result)
            return result
        except XlError as exc:
            raise SharedXlError(CoreXlError(exc.code)) from exc

    return evaluate


def xl_iferror(value: Callable[[], object], fallback: Callable[[], object]) -> object:
    """Evaluate IFERROR lazily with shared error-consumer semantics."""
    return _call_shared(_shared_iferror, _shared_thunk(value), _shared_thunk(fallback))


def xl_ifna(value: Callable[[], object], fallback: Callable[[], object]) -> object:
    """Evaluate IFNA lazily, catching only NA errors."""
    return _call_shared(_shared_ifna, _shared_thunk(value), _shared_thunk(fallback))


def xl_iserror(value: Callable[[], object]) -> object:
    """Inspect an expression for errors without propagating them."""
    return _call_shared(_shared_iserror, _shared_thunk(value))


def xl_isna(value: Callable[[], object]) -> object:
    """Inspect an expression for an NA error."""
    return _call_shared(_shared_isna, _shared_thunk(value))


def xl_isblank(value: Callable[[], object]) -> object:
    """Inspect an expression for a blank using shared semantics."""
    return _call_shared(_shared_isblank, _shared_thunk(value))


def xl_isnumber_lazy(value: Callable[[], object]) -> object:
    """Inspect a possibly failing expression for a numeric value."""
    return _call_shared(_shared_isnumber, _shared_thunk(value))


def xl_npv(*args: object) -> object:
    """Compute NPV with the shared financial implementation."""
    return _shared_value(_shared_npv, *args)


def xl_rank(*args: object) -> object:
    """Compute RANK with the shared statistical implementation."""
    return _shared_value(_shared_rank, *args)


def xl_large(*args: object) -> object:
    """Compute LARGE with the shared statistical implementation."""
    return _shared_value(_shared_large, *args)


def xl_stdev(*args: object) -> object:
    """Compute STDEV with the shared statistical implementation."""
    return _shared_value(_shared_stdev, *args)


def xl_countif(*args: object) -> object:
    """Compute COUNTIF with the shared criteria implementation."""
    return _shared_value(_shared_countif, *args)


def xl_round(*args: object) -> object:
    """Round numbers with the shared Excel implementation."""
    return _shared_value(_shared_round, *args)


def xl_rounddown(*args: object) -> object:
    """Round toward zero with the shared Excel implementation."""
    return _shared_value(_shared_rounddown, *args)


def xl_numbervalue(*args: object) -> object:
    """Parse numeric text with the shared Excel implementation."""
    return _shared_value(_shared_numbervalue, *args)


def xl_left(*args: object) -> object:
    """Extract leading text with the shared Excel implementation."""
    return _shared_value(_shared_left, *args)


def xl_hlookup(*args: object) -> object:
    """Look up a horizontal table with shared Excel semantics."""
    return _shared_value(_shared_hlookup, *args)


def xl_lookup(*args: object) -> object:
    """Look up a vector or table with shared Excel semantics."""
    return _shared_value(_shared_lookup, *args)


def xl_xlookup(*args: object) -> object:
    """Look up corresponding arrays with shared Excel semantics."""
    return _shared_value(_shared_xlookup, *args)


def xl_at(values: Sequence[T], index: object) -> T:
    """Return `values[index]` (0-based), raising `#VALUE!` when out of range.

    `index` is coerced with core `to_number` and truncated toward zero.
    """
    position = int(_as_number(index))
    if position < 0 or position >= len(values):
        raise XlError("#VALUE!")
    return values[position]


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
        if not isinstance(rows, Mapping) or not isinstance(cols, Mapping):
            raise TypeError("view selects rows and columns by field name together")
        return _product_view(
            values,
            cast(Mapping[str, Sequence[object]], rows),
            cast(Mapping[str, Sequence[object]], cols),
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


def xl_raise(code: str) -> NoReturn:
    """Raise `XlError(code)` from generated expression position."""
    raise XlError(code)


def require_aligned(*series: Sequence[object]) -> int:
    """Return the common length, or fail if any series has a different length."""
    if not series:
        raise ValueError("require_aligned expected at least one series")
    lengths = [len(item) for item in series]
    if len(set(lengths)) != 1:
        raise ValueError(f"misaligned series lengths: {lengths}")
    return lengths[0]


def require_length(values: Sequence[object], length: int) -> None:
    """Fail if `values` is not a catalog-order array of `length`."""
    actual = len(values)
    if actual != length:
        raise ValueError(f"expected length {length}, got {actual}")


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
    """Private demand-driven series; only completed tensors are published."""

    def __init__(
        self,
        series_id: str,
        domain: Domain,
        compute: Callable[[tuple[str | int, ...]], T],
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
        self._domain.require(coordinate)
        identity = (self._series_id, coordinate)
        if identity in self._memo:
            return self._memo[identity]
        if identity in self._active:
            raise InstanceCycleError(f"circular reference at {self._series_id}{coordinate!r}")
        self._active.add(identity)
        try:
            value = self._compute(coordinate)
            self._memo[identity] = value
            return value
        finally:
            self._active.remove(identity)


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


def live_measure(value: T) -> T:
    """Return `value`, or raise `XlError` when it is a stored error code."""
    if isinstance(value, str) and is_error(value):
        raise XlError(value)
    return value


def demand_instance(
    statement: str,
    index: int,
    compute: Callable[[int], T],
    memo: dict[tuple[str, int], T],
    stack: set[tuple[str, int]],
) -> T:
    """Like `eval_instance`, but re-raise a stored Excel error as `XlError`."""
    value = eval_instance(statement, index, compute, memo, stack)
    if isinstance(value, str) and is_error(value):
        raise XlError(value)
    return value
