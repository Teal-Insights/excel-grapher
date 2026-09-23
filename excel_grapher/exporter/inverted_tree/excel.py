"""Excel value semantics for inverted-tree codegen.

Wrappers adapt `core` / `export_runtime` sentinels into inverted-tree `XlError`
exceptions and stored error-code measures. Generated packages embed this module
as `excel.py`.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from datetime import date, datetime
from typing import Literal, NoReturn, TypeGuard, TypeVar, cast, overload

from excel_grapher.core.coercions import to_bool, to_number
from excel_grapher.core.grid import Range
from excel_grapher.core.logic_funcs import logical_and, logical_if, logical_not, logical_or
from excel_grapher.core.lookup_funcs import index_cells, match_cells, vlookup_cells
from excel_grapher.core.math_funcs import (
    abs_number,
    average_cells,
    exp_number,
    max_cells,
    min_cells,
    sum_cells,
)
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
from excel_grapher.core.text_funcs import value_from_text
from excel_grapher.core.types import CellValue, FormulaValue
from excel_grapher.core.types import XlError as CoreXlError
from excel_grapher.core.types import XlErrorException as SharedXlError
from excel_grapher.exporter.export_runtime.error_funcs import xl_iferror as _shared_iferror
from excel_grapher.exporter.export_runtime.error_funcs import xl_ifna as _shared_ifna
from excel_grapher.exporter.export_runtime.error_funcs import xl_isblank as _shared_isblank
from excel_grapher.exporter.export_runtime.error_funcs import xl_iserror as _shared_iserror
from excel_grapher.exporter.export_runtime.error_funcs import xl_isna as _shared_isna
from excel_grapher.exporter.export_runtime.error_funcs import xl_isnumber as _shared_isnumber
from excel_grapher.exporter.export_runtime.error_funcs import xl_istext as _shared_istext
from excel_grapher.exporter.export_runtime.info import xl_count as _shared_count
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
from excel_grapher.exporter.export_runtime.text import xl_left as _shared_left
from excel_grapher.exporter.export_runtime.text import xl_numbervalue as _shared_numbervalue
from excel_grapher.exporter.export_runtime.text import xl_text as _shared_text
from excel_grapher.series_bindings.input_coerce import (
    apply_input_value_map as apply_input_value_map,
)
from excel_grapher.series_bindings.input_coerce import (
    coerce_input_measure as coerce_input_measure,
)
from excel_grapher.series_bindings.input_coerce import require_input_domain as require_input_domain

T = TypeVar("T")

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
def as_measure(value: object, dtype: Literal["str"]) -> str | int | float | bool: ...


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
    A `str` measure keeps Excel numbers and bools so `INDEX(...)=1` on a
    string-dtyped computed 0/1 series stays numeric, matching the evaluator.
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
        if isinstance(value, bool | int | float):
            return value
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
    """Prepare an arithmetic operand for `core`.

    Blank cells (`None`) stay `None` so `to_number` coerces them to `0`. Empty
    text (`""`) is left as text so arithmetic raises `#VALUE!` (Excel / #420).
    """
    _raise_stored_error(value)
    return _as_formula(value)


def _as_number(value: object) -> float:
    """Coerce `value` via core `to_number`, re-raising stored error codes."""
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


def xl_bool(value: object) -> bool:
    """Coerce an `IF` condition with Excel boolean rules.

    `to_bool("")` is `False` (evaluator / `to_bool`), not Excel's `#VALUE!`.
    Non-boolean text such as `"nope"` raises `#VALUE!`.
    """
    _raise_stored_error(value)
    result = _adapt_core(to_bool(_as_formula(value)))
    assert isinstance(result, bool), f"IF condition returned {type(result).__name__}"
    return result


def xl_exp(*args: object) -> object:
    """Excel `EXP` via `core.math_funcs.exp_number`."""
    for arg in args:
        _raise_stored_error(arg)
    return _adapt_core(exp_number(*cast(tuple[CellValue, ...], args)))


def xl_abs(*args: object) -> object:
    """Excel `ABS` via `core.math_funcs.abs_number`."""
    for arg in args:
        _raise_stored_error(arg)
    return _adapt_core(abs_number(*cast(tuple[CellValue, ...], args)))


def xl_value(*args: object) -> object:
    """Excel `VALUE` via `core.text_funcs.value_from_text`."""
    for arg in args:
        _raise_stored_error(arg)
    if len(args) != 1:
        raise XlError("#VALUE!")
    return _adapt_core(value_from_text(cast(CellValue, args[0])))


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


def xl_lookup_cell(measure: object, workbook: object) -> object:
    """Return `measure`, restoring `workbook`'s Excel type after dtype stringify.

    INDEX/MATCH tables still read the bound series so `overrides` apply. A
    string measure that is only `str(workbook)` is the series dtype hiding a
    number or bool; Excel `INDEX` returns the worksheet type.
    """
    if measure == workbook:
        return measure
    if isinstance(measure, str) and measure == str(workbook):
        return workbook
    return measure


def _as_native_grid(natives: object, height: int, width: int) -> tuple[tuple[object, ...], ...]:
    """Interpret `natives` as a `height` by `width` row-major grid."""
    if isinstance(natives, str) or not isinstance(natives, Sequence):
        raise TypeError("natives must be a nested sequence")
    rows = list(natives)
    if height == 1 and width == len(rows) and (not rows or not isinstance(rows[0], Sequence)):
        return (tuple(rows),)
    grid: list[tuple[object, ...]] = []
    for row in rows:
        if isinstance(row, str) or not isinstance(row, Sequence):
            grid.append((row,))
        else:
            grid.append(tuple(row))
    if len(grid) != height or any(len(row) != width for row in grid):
        got = f"{len(grid)}x{len(grid[0]) if grid else 0}"
        raise ValueError(f"natives shape {got} != {height}x{width}")
    return tuple(grid)


def xl_typed_range(values: Range, natives: object) -> Range:
    """Apply `xl_lookup_cell` to each cell of `values` using `natives`.

    `natives` is a row-major nested sequence matching `values.shape`.
    """
    height, width = values.shape
    grid = _as_native_grid(natives, height, width)

    def resolve(row: int, column: int) -> FormulaValue:
        return cast(
            FormulaValue,
            xl_lookup_cell(values.cell(row, column), grid[row - 1][column - 1]),
        )

    return Range("", 1, 1, height, width, lambda address: None, _coord_resolver=resolve)


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


def xl_isnumber(value: Callable[[], object]) -> object:
    """Inspect a possibly failing expression for a numeric value."""
    return _call_shared(_shared_isnumber, _shared_thunk(value))


def xl_istext(value: Callable[[], object]) -> object:
    """Inspect a possibly failing expression for a text value."""
    return _call_shared(_shared_istext, _shared_thunk(value))


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


def xl_count(*args: object) -> object:
    """Count numeric values with the shared COUNT implementation."""
    return _shared_value(_shared_count, *args)


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


def xl_text(*args: object) -> object:
    """Format a value as text with the shared Excel implementation."""
    return _shared_value(_shared_text, *args)


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


def xl_raise(code: str) -> NoReturn:
    """Raise `XlError(code)` from generated expression position."""
    raise XlError(code)
