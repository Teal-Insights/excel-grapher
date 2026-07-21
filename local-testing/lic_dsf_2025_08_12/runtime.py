"""Standalone runtime for generated Excel formula code."""

from __future__ import annotations

import warnings
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, field
from datetime import date, datetime
from enum import StrEnum
from typing import NoReturn, TypeAlias, cast

import fastpyxl.utils.cell
import math
import numpy as np
import re

class CircularReferenceWarning(RuntimeWarning):
    """Warning emitted when a circular reference is encountered (default Excel mode)."""

@dataclass(slots=True)
class EvalContextBase:
    """Per-run evaluation state without dependency-tracking fields."""

    inputs: dict[str, CellValue]
    resolver: Callable[[str], Callable[[EvalContext], CellValue] | None]
    cache: dict[str, CellValue] = field(default_factory=dict)
    computing: set[str] = field(default_factory=set)
    circular_warning_roots: set[str] = field(default_factory=set)
    iterative_enabled: bool = False
    iterate_count: int = 100
    iterate_delta: float = 0.001
    iteration_values: dict[str, CellValue] = field(default_factory=dict)

EvalContext = EvalContextBase

NormalizedAddress: TypeAlias = str

class XlError(StrEnum):
    VALUE = "#VALUE!"
    REF = "#REF!"
    DIV = "#DIV/0!"
    NA = "#N/A"
    NAME = "#NAME?"
    NUM = "#NUM!"
    NULL = "#NULL!"

    @classmethod
    def from_text(cls, value: str) -> XlError | None:
        upper = value.strip().upper()
        for err in cls:
            if err.value == upper:
                return err
        return None

Scalar: TypeAlias = float | int | str | bool | XlError | None

class XlErrorException(Exception):
    """Exception form of an Excel error code.

    The exported runtime raises Excel errors as exceptions; the evaluator keeps
    `XlError` sentinel values and never raises this type.
    """

    code: XlError

    def __init__(self, code: XlError) -> None:
        """Initialize the exception with an Excel error code."""
        if not isinstance(code, XlError):
            raise TypeError(f"Expected XlError, got {type(code).__name__}")
        self.code = code
        super().__init__(code.value)

_EXCEL_EPOCH = datetime(1899, 12, 30)

def _criteria_compare(op: str, left: T, right: T) -> bool:
    """Compare two values of the same type."""
    if op == "=":
        return left == right
    if op == "<>":
        return left != right
    if op == ">":
        return left > right
    if op == "<":
        return left < right
    if op == ">=":
        return left >= right
    if op == "<=":
        return left <= right
    return False

def _escape_sheet_for_formula(sheet: str) -> str:
    """Escape apostrophes for use inside quoted sheet names."""
    return sheet.replace("'", "''")

def _format_general_number(value: float | int) -> str:
    f = float(value)
    if f.is_integer():
        return str(int(f))
    return str(f)

def _iter_numeric_cells(values: list[CellValue]) -> tuple[list[float], XlError | None]:
    nums: list[float] = []
    for v in values:
        if isinstance(v, XlError):
            return ([], v)
        if v is None:
            continue
        if isinstance(v, bool):
            continue
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            nums.append(float(v))
            continue
        if isinstance(v, (np.integer, np.floating)):
            nums.append(float(v))
            continue
    return (nums, None)

def _parse_countif_criteria(criteria: str) -> tuple[str | None, str]:
    s = criteria.strip()
    for op in (">=", "<=", "<>", ">", "<", "="):
        if s.startswith(op):
            return (op, s[len(op) :].strip())
    return (None, s)

def _raise_error(code: XlError) -> XlErrorException:
    """Build the exception for an Excel error code (callers raise the result)."""
    return XlErrorException(code)

def _cell_or_raise(grid: Grid, row0: int, col0: int) -> Scalar:
    """Read one grid cell, raising when the stored value is an error sentinel."""
    value = grid.at(row0, col0)
    if isinstance(value, XlError):
        raise _raise_error(value)
    return value

def _raise_if_error_value(value: CellValue) -> CellValue:
    """Surface Excel error values as raised exceptions at the cell boundary."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value

def _vector_of(grid: Grid) -> Grid | None:
    if grid.nrows == 1 or grid.ncols == 1:
        return grid
    return None

def _wildcard_to_regex(pattern: str) -> re.Pattern[str]:
    out: list[str] = ["^"]
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern):
            i += 1
            out.append(re.escape(pattern[i]))
        elif ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    out.append("$")
    return re.compile("".join(out), re.IGNORECASE)

def apply_arithmetic(op: str, ln: float, rn: float) -> float | XlError:
    """Apply an Excel arithmetic operator to two coerced numbers."""
    if op == "+":
        return ln + rn
    if op == "-":
        return ln - rn
    if op == "*":
        return ln * rn
    if op == "/":
        if rn == 0:
            return XlError.DIV
        return ln / rn
    if op == "^":
        try:
            value = ln**rn
        except (ValueError, OverflowError):
            return XlError.NUM
        if isinstance(value, complex):
            return XlError.NUM
        return value
    raise ValueError(f"Unknown arithmetic operator: {op}")

def datetime_to_excel_serial(value: datetime) -> float:
    """Convert a naive datetime to an Excel day serial (1900 date system)."""
    naive = value.replace(tzinfo=None) if value.tzinfo is not None else value
    delta = naive - _EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86_400.0

def _try_parse_iso_date_serial(text: str) -> float | None:
    stripped = text.strip()
    if not stripped:
        return None
    try:
        if "T" in stripped or " " in stripped:
            parsed = datetime.fromisoformat(stripped.replace("Z", "+00:00"))
            if parsed.tzinfo is not None:
                parsed = parsed.replace(tzinfo=None)
        else:
            parsed = datetime.combine(date.fromisoformat(stripped), datetime.min.time())
        return datetime_to_excel_serial(parsed)
    except ValueError:
        return None

def excel_casefold(value: str) -> str:
    return value.casefold()

def needs_quoting(sheet: str) -> bool:
    """Return True if a sheet name must be wrapped in single quotes in a formula."""
    return " " in sheet or "-" in sheet or "'" in sheet

def parse_address(address: str) -> tuple[str, str]:
    """Parse a sheet-qualified address into `(sheet, cell_coord)`.

    The returned sheet name has any surrounding single quotes stripped and any
    escaped apostrophes (`''`) unescaped to a single apostrophe.

    Examples:
        >>> parse_address("Sheet1!A1")
        ('Sheet1', 'A1')
        >>> parse_address("'My Sheet'!B2")
        ('My Sheet', 'B2')
        >>> parse_address("'It''s Data'!C3")
        ("It's Data", 'C3')
    """
    if address.startswith("'"):
        i = 1
        while i < len(address):
            if address[i] == "'":
                if i + 1 < len(address) and address[i + 1] == "'":
                    i += 2
                    continue
                break
            i += 1
        sheet = address[1:i].replace("''", "'")
        rest = address[i + 1 :]
        if rest.startswith("!"):
            return sheet, rest[1:]
        raise ValueError(f"Invalid address format: {address}")

    if "!" in address:
        sheet, cell = address.rsplit("!", 1)
        return sheet, cell

    raise ValueError(f"Address must be sheet-qualified: {address}")

def quote_sheet_if_needed(sheet: str) -> str:
    """Return a sheet name quoted for formulas when quoting is required."""
    if not needs_quoting(sheet):
        return sheet
    return "'" + _escape_sheet_for_formula(sheet) + "'"

def format_cell_key(sheet: str, column: str, row: int) -> NormalizedAddress:
    """Format a (sheet, column_letters, row) triple into a canonical address."""
    return f"{quote_sheet_if_needed(sheet)}!{column}{row}"

@dataclass(frozen=True, slots=True)
class Range:
    """Rectangular lazy range with consumer-driven cell access.

    Coordinates passed to `cell`, `row`, `column`, and `view` are 1-based and
    relative to this range, matching Excel function arguments.
    """

    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    # Resolvers may come from evaluation contexts with their own value
    # vocabulary; values are validated/coerced at consumption time.
    _resolver: Callable[[str], CellValue] = field(repr=False, compare=False)

    def __post_init__(self) -> None:
        """Validate the rectangular bounds."""
        if self.start_row < 1 or self.start_col < 1:
            raise ValueError("Range coordinates must be positive")
        if self.end_row < self.start_row or self.end_col < self.start_col:
            raise ValueError("Range end must be greater than or equal to start")

    @property
    def shape(self) -> tuple[int, int]:
        """The range shape as `(rows, columns)`."""
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        """Yield row-major addresses without evaluating cells."""
        for row in range(self.start_row, self.end_row + 1):
            for col in range(self.start_col, self.end_col + 1):
                yield self._address(row, col)

    def cell(self, row: int, col: int) -> CellValue:
        """Return a single relative cell value without evaluating siblings.

        Args:
            row: 1-based row within the range.
            col: 1-based column within the range.

        Raises:
            IndexError: If `row` or `col` is outside the range.
            XlErrorException: If the resolved cell is an Excel error.
        """
        self._validate_relative_cell(row, col)
        value = self._resolver(self._address(self.start_row + row - 1, self.start_col + col - 1))
        return self._raise_if_error(value)

    def row(self, row: int) -> Range:
        """Return a lazy view for one relative row."""
        nrows, _ = self.shape
        if row < 1 or row > nrows:
            raise IndexError("Range row is out of bounds")
        absolute_row = self.start_row + row - 1
        return Range(
            self.sheet,
            absolute_row,
            self.start_col,
            absolute_row,
            self.end_col,
            self._resolver,
        )

    def column(self, col: int) -> Range:
        """Return a lazy view for one relative column."""
        _, ncols = self.shape
        if col < 1 or col > ncols:
            raise IndexError("Range column is out of bounds")
        absolute_col = self.start_col + col - 1
        return Range(
            self.sheet,
            self.start_row,
            absolute_col,
            self.end_row,
            absolute_col,
            self._resolver,
        )

    def view(
        self,
        row_start: int = 1,
        row_end: int | None = None,
        col_start: int = 1,
        col_end: int | None = None,
    ) -> Range:
        """Return a lazy rectangular subrange view using relative coordinates."""
        nrows, ncols = self.shape
        row_end = nrows if row_end is None else row_end
        col_end = ncols if col_end is None else col_end
        self._validate_relative_cell(row_start, col_start)
        self._validate_relative_cell(row_end, col_end)
        if row_end < row_start or col_end < col_start:
            raise ValueError("Range view end must be greater than or equal to start")
        return Range(
            self.sheet,
            self.start_row + row_start - 1,
            self.start_col + col_start - 1,
            self.start_row + row_end - 1,
            self.start_col + col_end - 1,
            self._resolver,
        )

    def value_at(self, row: int, col: int) -> CellValue:
        """Return a single relative cell value with errors as sentinels.

        Unlike `cell`, Excel errors surface as `XlError` sentinel values (raised
        `XlErrorException`s from the resolver are caught and converted). Range
        consumers that implement Excel skip semantics (lookup scans, criteria
        matching) use this accessor; `cell`/iteration raise instead.
        """
        self._validate_relative_cell(row, col)
        address = self._address(self.start_row + row - 1, self.start_col + col - 1)
        try:
            return self._resolver(address)
        except XlErrorException as exc:
            return exc.code

    def iter_raw(self) -> Iterator[CellValue]:
        """Yield raw values (error sentinels included) in row-major order."""
        nrows, ncols = self.shape
        for row in range(1, nrows + 1):
            for col in range(1, ncols + 1):
                yield self.value_at(row, col)

    def rows_raw(self) -> list[list[CellValue]]:
        """Materialize the range as nested row lists of raw values."""
        nrows, ncols = self.shape
        return [[self.value_at(r, c) for c in range(1, ncols + 1)] for r in range(1, nrows + 1)]

    def iter_values(self) -> Iterator[CellValue]:
        """Yield values in deterministic row-major order."""
        nrows, ncols = self.shape
        for row in range(1, nrows + 1):
            for col in range(1, ncols + 1):
                yield self.cell(row, col)

    def __iter__(self) -> Iterator[CellValue]:
        """Yield values in deterministic row-major order."""
        return self.iter_values()

    def _address(self, row: int, col: int) -> str:
        col_letter = fastpyxl.utils.cell.get_column_letter(col)
        return format_cell_key(self.sheet, col_letter, row)

    def _validate_relative_cell(self, row: int, col: int) -> None:
        nrows, ncols = self.shape
        if row < 1 or row > nrows or col < 1 or col > ncols:
            raise IndexError("Range cell is out of bounds")

    @staticmethod
    def _raise_if_error(value: CellValue) -> CellValue:
        if isinstance(value, XlError):
            raise XlErrorException(value)
        return value

class Grid:
    """Positional raw-value access over a lazy `Range` or nested-list array."""

    __slots__ = ("nrows", "ncols", "_range", "_rows")

    def __init__(
        self,
        nrows: int,
        ncols: int,
        rng: Range | None,
        rows: list[list[CellValue]] | None,
    ) -> None:
        self.nrows = nrows
        self.ncols = ncols
        self._range = rng
        self._rows = rows

    @staticmethod
    def wrap(value: object) -> Grid | None:
        """Wrap a range/array value; return `None` for scalar values."""
        if isinstance(value, Range):
            nrows, ncols = value.shape
            return Grid(nrows, ncols, value, None)
        if isinstance(value, (list, tuple)):
            rows = [
                list(row) if isinstance(row, (list, tuple)) else [row]
                for row in cast("list[CellValue]", value)
            ]
            if not rows:
                rows = [[None]]
            return Grid(len(rows), len(rows[0]), None, cast("list[list[CellValue]]", rows))
        return None

    def at(self, row0: int, col0: int) -> Scalar:
        """Return the raw value at a 0-based position (error sentinels included)."""
        if self._range is not None:
            return cast(Scalar, self._range.value_at(row0 + 1, col0 + 1))
        assert self._rows is not None
        return cast(Scalar, self._rows[row0][col0])

    def at_flat(self, index0: int) -> Scalar:
        """Return the raw value at a 0-based row-major flat index."""
        row0, col0 = divmod(index0, self.ncols)
        return self.at(row0, col0)

    @property
    def size(self) -> int:
        """Total cell count."""
        return self.nrows * self.ncols

    def iter_raw(self) -> Iterator[Scalar]:
        """Yield raw values (error sentinels included) in row-major order."""
        for row0 in range(self.nrows):
            for col0 in range(self.ncols):
                yield self.at(row0, col0)

    def row_slice(self, row0: int) -> Range | list[list[CellValue]]:
        """Return one row as a lazy view (`Range` input) or nested list."""
        if self._range is not None:
            return self._range.row(row0 + 1)
        assert self._rows is not None
        return [list(self._rows[row0])]

    def col_slice(self, col0: int) -> Range | list[list[CellValue]]:
        """Return one column as a lazy view (`Range` input) or nested list."""
        if self._range is not None:
            return self._range.column(col0 + 1)
        assert self._rows is not None
        return [[row[col0]] for row in self._rows]

def _as_scalar(value: object) -> Scalar:
    if isinstance(value, (Range, list, tuple)):
        return XlError.VALUE
    if Grid.wrap(value) is not None:
        return XlError.VALUE
    return cast(Scalar, value)

def _broadcast_pair(left: CellValue, right: CellValue) -> tuple[Grid, Grid] | None:
    """Wrap array operands as aligned grids; `None` when both operands are scalar."""
    left_grid = Grid.wrap(left)
    right_grid = Grid.wrap(right)
    if left_grid is None and right_grid is None:
        return None
    if left_grid is not None and right_grid is not None:
        if (left_grid.nrows, left_grid.ncols) != (right_grid.nrows, right_grid.ncols):
            raise _raise_error(XlError.VALUE)
        return left_grid, right_grid
    if left_grid is not None:
        scalar_right = Grid.wrap([[right] * left_grid.ncols for _ in range(left_grid.nrows)])
        assert scalar_right is not None
        return left_grid, scalar_right
    assert right_grid is not None
    scalar_left = Grid.wrap([[left] * right_grid.ncols for _ in range(right_grid.nrows)])
    assert scalar_left is not None
    return scalar_left, right_grid

def _coerce_grid(value: object) -> Grid | XlError | None:
    """Wrap array-like values; return `None` for scalars, errors as-is."""
    if isinstance(value, XlError):
        return value
    return Grid.wrap(value)

def _format_address(sheet: str, row: int, col: int) -> str:
    return format_cell_key(sheet, fastpyxl.utils.cell.get_column_letter(col), row)

def _resolve_scalar(value_fn: Callable[[], CellValue]) -> CellValue:
    """Evaluate a thunk, resolving 1x1 range views to their single cell value."""
    value = value_fn()
    if isinstance(value, Range) and value.shape == (1, 1):
        return cast("CellValue", value.cell(1, 1))
    return value

def flatten(*args: CellValue) -> Iterator[Scalar]:
    """Yield scalar values from scalars, nested lists, and lazy ranges.

    Excel errors raise `XlErrorException` on encounter, mirroring the
    evaluator's argument error precheck for generic worksheet functions.
    Lookup scans keep skip semantics through `Grid`/`Range.value_at` instead.
    """
    for arg in args:
        if isinstance(arg, Range):
            for v in arg.iter_raw():
                if isinstance(v, XlError):
                    raise XlErrorException(v)
                yield cast(Scalar, v)
        elif isinstance(arg, (list, tuple)):
            yield from flatten(*arg)
        else:
            if isinstance(arg, XlError):
                raise XlErrorException(arg)
            yield cast(Scalar, arg)

def format_key(sheet: str, cell: str) -> NormalizedAddress:
    """Format a sheet and A1 cell coordinate into a canonical address string."""
    return f"{quote_sheet_if_needed(sheet)}!{cell}"

@dataclass(frozen=True, slots=True)
class ExcelRange:
    """Rectangular worksheet reference geometry for evaluator and export."""

    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def shape(self) -> tuple[int, int]:
        """The reference shape as `(rows, columns)`."""
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        """Yield row-major sheet-qualified addresses without evaluating cells."""
        for r in range(self.start_row, self.end_row + 1):
            for c in range(self.start_col, self.end_col + 1):
                col = fastpyxl.utils.cell.get_column_letter(c)
                yield format_key(self.sheet, f"{col}{r}")

CellValue: TypeAlias = Scalar | ExcelRange | Range | list["CellValue"]

def _raise_if_error(value: object) -> CellValue:
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return cast(CellValue, value)

def _range_from_ref_info(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
) -> ExcelRange:
    """Normalize generated reference metadata into an `ExcelRange`."""
    if isinstance(ref, ExcelRange):
        return ref
    match ref:
        case (sheet, base_row, base_col):
            return ExcelRange(
                sheet=sheet,
                start_row=base_row,
                start_col=base_col,
                end_row=base_row,
                end_col=base_col,
            )
        case (sheet, base_row, base_col, base_end_row, base_end_col):
            return ExcelRange(
                sheet=sheet,
                start_row=base_row,
                start_col=base_col,
                end_row=base_end_row,
                end_col=base_end_col,
            )
        case _:
            raise XlErrorException(XlError.VALUE)

def as_scalar(value: CellValue) -> Scalar:
    """Collapse range/array values to `#VALUE!` for scalar coercion contexts.

    Keep behavior aligned with `excel_grapher.core.coercions.as_scalar`. This
    module is embedded into standalone exports and cannot import library code.
    """
    if isinstance(value, (Range, ExcelRange, list, tuple)):
        return XlError.VALUE
    return value

def _as_addressing_scalar(value: CellValue | None) -> Scalar | None:
    """Collapse export-runtime values to scalars for shared addressing helpers."""
    if value is None:
        return None
    return as_scalar(value)

def _scalar_or_raise(value: CellValue) -> Scalar:
    """Collapse to a scalar, raising when the value is an Excel error."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    return scalar

def coerce_inputs_dict(values: Mapping[str, object]) -> dict[str, CellValue]:
    """Widen inferred default-input dicts to `dict[str, CellValue]` for `EvalContext`."""
    return cast(dict[str, CellValue], dict(values))

def columns_count(ref: CellValue) -> int | XlError:
    """Return the column count of a reference."""
    if isinstance(ref, ExcelRange):
        return ref.end_col - ref.start_col + 1
    return XlError.VALUE

def get_error(*args: object) -> XlError | None:
    """Return the first top-level / flattened-list `XlError`, if any.

    Does not evaluate cells inside a lazy `Range` (see `flatten`).
    """
    for v in flatten(*args):
        if isinstance(v, XlError):
            return v
    return None

def raise_if_sentinel_bool(value: bool | XlError) -> bool:
    """Return a boolean result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value

def raise_if_sentinel_float(value: float | XlError) -> float:
    """Return a float result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value

def raise_if_sentinel_int(value: int | XlError) -> int:
    """Return an integer result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value

def raise_if_sentinel_str(value: str | XlError) -> str:
    """Return a string result or raise ``XlErrorException`` for an error sentinel."""
    if isinstance(value, XlError):
        raise XlErrorException(value)
    return value

def row_number(ref: CellValue) -> int | XlError:
    """Return the row number of a reference."""
    if isinstance(ref, ExcelRange):
        return ref.start_row
    return XlError.VALUE

def split_sheet_qualified_address(address: str) -> tuple[str, str] | None:
    """Split `sheet!coord` into `(sheet_name, coord)`.

    Handles quoted sheet names, including Excel's doubled-single-quote escape
    (`'O''Neil'!A1` -> sheet `O'Neil`).

    Returns `None` when *address* has no sheet qualifier (plain `A1`).
    """
    if "!" not in address:
        return None
    try:
        return parse_address(address)
    except ValueError:
        return None

def _parse_sheet_address(address: str) -> tuple[str, str] | None:
    return split_sheet_qualified_address(address)

def _parse_range_address(address: str) -> tuple[str, str, str] | XlError:
    if ":" not in address:
        return XlError.VALUE
    start_text, end_text = address.split(":", 1)
    start = _parse_sheet_address(start_text)
    if start is None:
        return XlError.VALUE
    sheet, start_cell = start
    if "!" in end_text:
        end = _parse_sheet_address(end_text)
        if end is None:
            return XlError.VALUE
        end_sheet, end_cell = end
        if end_sheet != sheet:
            return XlError.VALUE
    else:
        end_cell = end_text
    return sheet, start_cell, end_cell

def to_bool(value: CellValue) -> bool | XlError:
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        return scalar
    value = cast(CellValue, scalar)
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0
    if isinstance(value, str):
        s = value.strip().upper()
        if s == "":
            return False
        if s == "TRUE":
            return True
        if s == "FALSE":
            return False
        return XlError.VALUE
    return XlError.VALUE

def logical_and(*args: CellValue) -> bool | XlError:
    """Return logical AND across arguments."""
    err = get_error(*args)
    if err is not None:
        return err
    for a in args:
        b = to_bool(a)
        if isinstance(b, XlError):
            return b
        if not b:
            return False
    return True

def logical_or(*args: CellValue) -> bool | XlError:
    """Return logical OR across arguments."""
    err = get_error(*args)
    if err is not None:
        return err
    for a in args:
        b = to_bool(a)
        if isinstance(b, XlError):
            return b
        if b:
            return True
    return False

def to_string(value: CellValue) -> str:
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        return scalar.value
    value = cast(CellValue, scalar)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return _format_general_number(float(value))
    if isinstance(value, str):
        return value
    return str(value)

def concat_scalars(left: CellValue, right: CellValue) -> str:
    return to_string(left) + to_string(right)

def try_coerce_string_to_float(text: str) -> float | None:
    """Parse one Excel numeric string, or return None when coercion fails."""
    stripped = text.strip()
    if stripped == "":
        return 0.0
    try:
        return float(stripped)
    except ValueError:
        return _try_parse_iso_date_serial(stripped)

def to_number(value: CellValue) -> float | XlError:
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        return scalar
    value = cast(CellValue, scalar)
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        number = try_coerce_string_to_float(value)
        if number is None:
            return XlError.VALUE
        return number
    return XlError.VALUE

def _compare_values(a: object, b: object) -> int:
    a = _as_scalar(a)
    b = _as_scalar(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return -1 if an < bn else 1 if an > bn else 0
    if isinstance(a, str) and isinstance(b, str):
        af = excel_casefold(a)
        bf = excel_casefold(b)
        return -1 if af < bf else 1 if af > bf else 0
    return 0

def _number_or_raise(value: CellValue) -> float:
    """Coerce a scalar argument to a number, raising on Excel coercion errors."""
    number = to_number(as_scalar(value))
    if isinstance(number, XlError):
        raise XlErrorException(number)
    return number

def _value_matches_criteria(cell_value: CellValue, criteria: CellValue) -> bool:
    if isinstance(criteria, XlError):
        return False
    if not isinstance(criteria, str):
        target = criteria
        if isinstance(cell_value, XlError):
            return False
        if target is None:
            return cell_value is None
        if isinstance(target, bool):
            b = to_bool(cell_value)
            return (not isinstance(b, XlError)) and b == target
        if isinstance(target, (int, float)) and not isinstance(target, bool):
            vn = to_number(cell_value)
            return (not isinstance(vn, XlError)) and vn == float(target)
        return excel_casefold(to_string(cell_value)) == excel_casefold(to_string(target))

    op, rhs = _parse_countif_criteria(criteria)
    if isinstance(cell_value, XlError):
        return False

    if op is None:
        if any(ch in rhs for ch in ("*", "?", "~")):
            rx = _wildcard_to_regex(rhs)
            return rx.match(to_string(cell_value)) is not None
        return excel_casefold(to_string(cell_value)) == excel_casefold(rhs)

    try:
        rhs_num = float(rhs) if rhs != "" else 0.0
    except ValueError:
        rhs_num = None

    if rhs_num is not None:
        vn = to_number(cell_value)
        if isinstance(vn, XlError):
            return False
        return _criteria_compare(op, vn, rhs_num)

    return _criteria_compare(op, excel_casefold(to_string(cell_value)), excel_casefold(rhs))

def _values_match(a: object, b: object) -> bool:
    a = _as_scalar(a)
    b = _as_scalar(b)
    if isinstance(a, str) and isinstance(b, str):
        return excel_casefold(a) == excel_casefold(b)
    an = to_number(a)
    bn = to_number(b)
    if not isinstance(an, XlError) and not isinstance(bn, XlError):
        return an == bn
    return a == b

def compare_scalars(op: str, left: CellValue, right: CellValue) -> bool | XlError:
    """Compare two scalar cell values using Excel coercion rules."""
    if isinstance(left, XlError):
        return left
    if isinstance(right, XlError):
        return right

    def _cmp_str(a: str, b: str) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    def _cmp_float(a: float, b: float) -> bool:
        if op == "=":
            return a == b
        if op == "<>":
            return a != b
        if op == "<":
            return a < b
        if op == ">":
            return a > b
        if op == "<=":
            return a <= b
        if op == ">=":
            return a >= b
        raise ValueError(f"Unknown comparison operator: {op}")

    if isinstance(left, str) and isinstance(right, str):
        return _cmp_str(excel_casefold(left), excel_casefold(right))

    ln = to_number(left)
    rn = to_number(right)
    if isinstance(ln, XlError) or isinstance(rn, XlError):
        return _cmp_str(excel_casefold(to_string(left)), excel_casefold(to_string(right)))

    return _cmp_float(float(ln), float(rn))

def countif_cells(range_values: CellValue, criteria: CellValue) -> int | XlError:
    """Count cells matching criteria."""
    if isinstance(criteria, XlError):
        return criteria
    values = list(flatten(range_values))
    return sum(1 for v in values if _value_matches_criteria(v, criteria))

def hlookup_cells(
    lookup_value: object,
    table_array: object,
    row_index_num: object,
    range_lookup: object = True,
) -> Scalar:
    """Excel HLOOKUP over a lazy grid or nested-list array."""
    rn = to_number(cast(CellValue, row_index_num))
    if isinstance(rn, XlError):
        return rn
    row_index = int(rn)
    if row_index < 1:
        return XlError.VALUE
    grid = Grid.wrap(table_array)
    if grid is None:
        return XlError.VALUE
    if row_index > grid.nrows:
        return XlError.REF
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.ncols):
            if _values_match(lookup_value, grid.at(0, i)):
                return grid.at(row_index - 1, i)
        return XlError.NA
    last_match_idx = None
    for i in range(grid.ncols):
        if _compare_values(grid.at(0, i), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return grid.at(row_index - 1, last_match_idx)

def index_cells(
    array: object,
    row_num: object = None,
    col_num: object = None,
) -> object:
    """Excel INDEX over a lazy grid or nested-list array.

    Returns a scalar cell value, or a row/column slice (`Range` or nested list)
    when only one of `row_num` / `col_num` selects a vector.
    """
    grid = Grid.wrap(array)
    if grid is None:
        return XlError.VALUE
    nrows, ncols = grid.nrows, grid.ncols
    row_omitted = row_num is None
    col_omitted = col_num is None

    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return grid.at(0, 0)
        if nrows == 1:
            return grid.at(0, ncols - 1)
        if ncols == 1:
            return grid.at(nrows - 1, 0)
        return XlError.VALUE

    if row_omitted:
        col_s = as_scalar(col_num)
        if isinstance(col_s, XlError):
            return col_s
        cn = to_number(cast(CellValue, col_s))
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return grid.at(0, col - 1)
        return grid.col_slice(col - 1)

    row_s = as_scalar(row_num)
    if isinstance(row_s, XlError):
        return row_s
    rn = to_number(cast(CellValue, row_s))
    if isinstance(rn, XlError):
        return rn
    row = int(rn)

    if col_omitted:
        if nrows == 1:
            if row < 1 or row > ncols:
                return XlError.REF
            return grid.at(0, row - 1)
        if ncols == 1:
            if row < 1 or row > nrows:
                return XlError.REF
            return grid.at(row - 1, 0)
        if row < 1 or row > nrows:
            return XlError.REF
        return grid.row_slice(row - 1)

    col_s = as_scalar(col_num)
    if isinstance(col_s, XlError):
        return col_s
    cn = to_number(cast(CellValue, col_s))
    if isinstance(cn, XlError):
        return cn
    col = int(cn)
    if nrows == 1:
        if row < 1 or row > ncols:
            return XlError.REF
        return grid.at(0, row - 1)
    if ncols == 1:
        if row < 1 or row > nrows:
            return XlError.REF
        return grid.at(row - 1, 0)
    if row < 1 or row > nrows:
        return XlError.REF
    if col < 1 or col > ncols:
        return XlError.REF
    return grid.at(row - 1, col - 1)

def index_excel_range(
    base: ExcelRangeGeometry,
    row_num: CellValue | None,
    col_num: CellValue | None,
) -> ExcelRange | XlError:
    """Map INDEX(row,col) over *base* to an absolute range (single cell or slice).

    Mirrors `excel_grapher.runtime.lookup.xl_index` geometry
    so OFFSET(INDEX(...), ...) receives a true cell reference.
    """
    nrows = base.end_row - base.start_row + 1
    ncols = base.end_col - base.start_col + 1
    row_omitted = row_num is None
    col_omitted = col_num is None

    def abs_cell(r0: int, c0: int) -> ExcelRange:
        r = base.start_row + r0
        c = base.start_col + c0
        return ExcelRange(base.sheet, r, c, r, c)

    if row_omitted and col_omitted:
        if nrows == 1 and ncols == 1:
            return abs_cell(0, 0)
        if nrows == 1:
            return abs_cell(0, ncols - 1)
        if ncols == 1:
            return abs_cell(nrows - 1, 0)
        return XlError.VALUE

    if row_omitted:
        cn = to_number(col_num)
        if isinstance(cn, XlError):
            return cn
        col = int(cn)
        if col < 1 or col > ncols:
            return XlError.REF
        if nrows == 1:
            return abs_cell(0, col - 1)
        c0 = base.start_col + col - 1
        return ExcelRange(base.sheet, base.start_row, c0, base.end_row, c0)

    rn = to_number(row_num)
    if isinstance(rn, XlError):
        return rn
    row = int(rn)

    if col_omitted:
        if nrows == 1:
            if row < 1 or row > ncols:
                return XlError.REF
            return abs_cell(0, row - 1)
        if ncols == 1:
            if row < 1 or row > nrows:
                return XlError.REF
            return abs_cell(row - 1, 0)
        if row < 1 or row > nrows:
            return XlError.REF
        r0 = base.start_row + row - 1
        return ExcelRange(base.sheet, r0, base.start_col, r0, base.end_col)

    cn = to_number(col_num)
    if isinstance(cn, XlError):
        return cn
    col = int(cn)
    if nrows == 1:
        if row < 1 or row > ncols:
            return XlError.REF
        return abs_cell(0, row - 1)
    if ncols == 1:
        if row < 1 or row > nrows:
            return XlError.REF
        return abs_cell(row - 1, 0)
    if row < 1 or row > nrows:
        return XlError.REF
    if col < 1 or col > ncols:
        return XlError.REF
    return abs_cell(row - 1, col - 1)

def large_kth(array: CellValue, k: CellValue) -> float | XlError:
    """Return the k-th largest numeric value."""
    kk = to_number(k)
    if isinstance(kk, XlError):
        return kk
    kth = int(kk)
    if kth < 1:
        return XlError.NUM
    values = list(flatten(array))
    nums, err = _iter_numeric_cells(values)
    if err is not None:
        return err
    if kth > len(nums):
        return XlError.NUM
    nums.sort(reverse=True)
    return float(nums[kth - 1])

def left_chars(text: CellValue, num_chars: CellValue = 1) -> str | XlError:
    """Return the leftmost characters of text."""
    scalar = as_scalar(text)
    if isinstance(scalar, XlError):
        return scalar
    s = to_string(cast(CellValue, scalar))
    n = to_number(num_chars)
    if isinstance(n, XlError):
        return n
    chars = int(n)
    if chars < 0:
        return XlError.VALUE
    return s[:chars]

def lookup_cells(
    lookup_value: object,
    lookup_vector_or_array: object,
    result_vector: object = None,
) -> Scalar:
    """Excel LOOKUP over a lazy grid or nested-list array."""
    grid = _coerce_grid(lookup_vector_or_array)
    if grid is None:
        return XlError.VALUE
    if isinstance(grid, XlError):
        return grid
    result_grid = _coerce_grid(result_vector) if result_vector is not None else None
    if result_vector is not None and result_grid is None:
        return XlError.VALUE
    if isinstance(result_grid, XlError):
        return result_grid

    if result_grid is None:
        vector = _vector_of(grid)
        if vector is not None:
            lookup_flat = vector
            result_flat = vector
        elif grid.nrows >= grid.ncols:
            lookup_flat = Grid.wrap(grid.col_slice(0))
            result_flat = Grid.wrap(grid.col_slice(grid.ncols - 1))
            assert lookup_flat is not None and result_flat is not None
        else:
            lookup_flat = Grid.wrap(grid.row_slice(0))
            result_flat = Grid.wrap(grid.row_slice(grid.nrows - 1))
            assert lookup_flat is not None and result_flat is not None
    else:
        if _vector_of(grid) is None or _vector_of(result_grid) is None:
            return XlError.NA
        if grid.size != result_grid.size:
            return XlError.NA
        lookup_flat = grid
        result_flat = result_grid

    last_match_idx = None
    for i in range(lookup_flat.size):
        if _compare_values(lookup_flat.at_flat(i), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return result_flat.at_flat(last_match_idx)

def match_cells(
    lookup_value: object,
    lookup_array: object,
    match_type: object = 1,
) -> int | XlError:
    """Excel MATCH over a lazy grid or nested-list array."""
    mt = to_number(cast(CellValue, match_type))
    if isinstance(mt, XlError):
        return mt
    match_type_int = int(mt)
    if isinstance(lookup_array, XlError):
        return lookup_array
    grid = Grid.wrap(lookup_array)
    if grid is None:
        grid_wrapped = Grid.wrap([[lookup_array]])
        assert grid_wrapped is not None
        grid = grid_wrapped
    if match_type_int == 0:
        for i in range(grid.size):
            if _values_match(lookup_value, grid.at_flat(i)):
                return i + 1
        return XlError.NA
    if match_type_int == 1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) <= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    if match_type_int == -1:
        last_match = None
        for i in range(grid.size):
            if _compare_values(grid.at_flat(i), lookup_value) >= 0:
                last_match = i + 1
            else:
                break
        return XlError.NA if last_match is None else last_match
    return XlError.VALUE

def numbervalue_parse(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float | XlError:
    """Convert text to a number with explicit decimal and group separators."""
    if isinstance(text, XlError):
        return text
    if isinstance(decimal_separator, XlError):
        return decimal_separator
    if isinstance(group_separator, XlError):
        return group_separator

    if not isinstance(text, str):
        return to_number(text)

    dec_sep = to_string(decimal_separator)
    grp_sep = to_string(group_separator)
    if dec_sep == "" or dec_sep == grp_sep:
        return XlError.VALUE

    s = text.replace("\u00a0", " ").strip()
    if s == "":
        return 0.0
    currency_symbols = "$€£¥"
    while s and (s[0] in currency_symbols or s[-1] in currency_symbols):
        s = s.lstrip(currency_symbols).rstrip(currency_symbols).strip()
        if s == "":
            return XlError.VALUE
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()
        if s == "":
            return XlError.VALUE
    percent = False
    if s.endswith("%"):
        percent = True
        s = s[:-1].strip()
        if s == "":
            return XlError.VALUE
    sign = 1.0
    if s.startswith(("+", "-")):
        if s[0] == "-":
            sign = -1.0
        s = s[1:].strip()
        if s == "":
            return XlError.VALUE
    while s and (s[0] in currency_symbols or s[-1] in currency_symbols):
        s = s.lstrip(currency_symbols).rstrip(currency_symbols).strip()
        if s == "":
            return XlError.VALUE
    if grp_sep:
        s = s.replace(grp_sep, "")
    if dec_sep != ".":
        s = s.replace(dec_sep, ".")
    try:
        value = float(s)
    except ValueError:
        return XlError.VALUE
    if percent:
        value /= 100.0
    if negative:
        value = -abs(value)
    return value * sign

def numeric_values(values: Iterable[CellValue]) -> tuple[list[float], XlError | None]:
    nums: list[float] = []
    for v in values:
        n = to_number(v)
        if isinstance(n, XlError):
            return ([], n)
        nums.append(float(n))
    return (nums, None)

def average_cells(*args: CellValue) -> float | XlError:
    """Return the average of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return XlError.DIV
    return float(sum(nums) / len(nums))

def max_cells(*args: CellValue) -> float | XlError:
    """Return the maximum of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(max(nums))

def min_cells(*args: CellValue) -> float | XlError:
    """Return the minimum of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) == 0:
        return 0.0
    return float(min(nums))

def npv_cells(rate: CellValue, *values: CellValue) -> float | XlError:
    """Return net present value for a rate and cash flows."""
    r = to_number(rate)
    if isinstance(r, XlError):
        return r
    all_values = list(flatten(*values))
    nums, err = numeric_values(all_values)
    if err is not None:
        return err
    if len(nums) == 0:
        return XlError.VALUE
    result = 0.0
    for i, val in enumerate(nums):
        result += val / ((1 + r) ** (i + 1))
    return result

def rank_number(number: CellValue, ref: CellValue, order: CellValue = 0) -> int | XlError:
    """Return the rank of a number within a reference range."""
    nn = to_number(number)
    if isinstance(nn, XlError):
        return nn
    oo = to_number(order)
    if isinstance(oo, XlError):
        return oo
    ascending = int(oo) != 0
    values = list(flatten(ref))
    nums, err = _iter_numeric_cells(values)
    if err is not None:
        return err
    if ascending:
        return 1 + sum(1 for v in nums if v < nn)
    return 1 + sum(1 for v in nums if v > nn)

def round_number(number: CellValue, num_digits: CellValue) -> float | XlError:
    """Round a number to the given number of digits."""
    n = to_number(number)
    if isinstance(n, XlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, XlError):
        return d
    return float(round(n, int(d)))

def rounddown_number(number: CellValue, num_digits: CellValue) -> float | XlError:
    """Round a number down to the given number of digits."""
    n = to_number(number)
    if isinstance(n, XlError):
        return n
    d = to_number(num_digits)
    if isinstance(d, XlError):
        return d
    digits = int(d)
    factor = 10**digits
    if n >= 0:
        return float(math.floor(n * factor) / factor)
    return float(math.ceil(n * factor) / factor)

def stdev_cells(*args: CellValue) -> float | XlError:
    """Return sample standard deviation of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    if len(nums) < 2:
        return XlError.DIV
    mean = sum(nums) / len(nums)
    variance = sum((x - mean) ** 2 for x in nums) / (len(nums) - 1)
    return float(variance**0.5)

def sum_cells(*args: CellValue) -> float | XlError:
    """Return the sum of numeric cells."""
    values = list(flatten(*args))
    nums, err = numeric_values(values)
    if err is not None:
        return err
    return float(sum(nums))

def to_int(value: CellValue) -> int | XlError:
    """Coerce a CellValue to an integer using Excel-style numeric coercion.

    For functions that operate on integer indices (e.g. CHOOSE/INDEX/MATCH)
    while propagating Excel errors.
    """
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return int(n)

def vlookup_cells(
    lookup_value: object,
    table_array: object,
    col_index_num: object,
    range_lookup: object = True,
) -> Scalar:
    """Excel VLOOKUP over a lazy grid or nested-list array."""
    cn = to_number(cast(CellValue, col_index_num))
    if isinstance(cn, XlError):
        return cn
    col_index = int(cn)
    if col_index < 1:
        return XlError.VALUE
    grid = Grid.wrap(table_array)
    if grid is None:
        return XlError.VALUE
    if col_index > grid.ncols:
        return XlError.REF
    exact_match = not bool(range_lookup)
    if exact_match:
        for i in range(grid.nrows):
            if _values_match(lookup_value, grid.at(i, 0)):
                return grid.at(i, col_index - 1)
        return XlError.NA
    last_match_idx = None
    for i in range(grid.nrows):
        if _compare_values(grid.at(i, 0), lookup_value) <= 0:
            last_match_idx = i
        else:
            break
    if last_match_idx is None:
        return XlError.NA
    return grid.at(last_match_idx, col_index - 1)

def warn_circular_reference(*, stacklevel: int = 2) -> None:
    """Emit the standard circular-reference warning."""
    warnings.warn(
        "Circular reference detected; returning 0 (iterative calculation is disabled).",
        CircularReferenceWarning,
        stacklevel=stacklevel,
    )

def xl_and(*args: CellValue) -> bool:
    """Return logical AND, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_and(*args))

def xl_average(*args: CellValue) -> float:
    """Return the average of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(average_cells(*args))

def xl_bool(value: CellValue) -> bool:
    """Coerce a scalar cell value to a boolean, raising on Excel errors."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    boolean = to_bool(scalar)
    if isinstance(boolean, XlError):
        raise _raise_error(boolean)
    return boolean

def xl_circular_reference() -> CellValue:
    """Excel default behavior for circular references (non-iterative calculation)."""
    warn_circular_reference(stacklevel=2)
    return 0

def _evaluate_address(
    ctx: EvalContextBase,
    address: str,
    obtain_fn: Callable[[], Callable[[EvalContextBase], CellValue]],
    *,
    preserve_structural_blank: bool = False,
) -> CellValue:
    """Shared evaluation path without dependency recording (slim export scaffold).

    Excel error values raise `XlErrorException`; the raising cell's error code
    is cached so re-reads raise without re-evaluating.
    """
    if address in ctx.cache:
        return _raise_if_error_value(ctx.cache[address])

    if address in ctx.computing:
        if ctx.iterative_enabled:
            return ctx.iteration_values.get(address, 0)
        return xl_circular_reference()

    if address in ctx.inputs:
        value = ctx.inputs[address]
        ctx.cache[address] = value
        return _raise_if_error_value(value)

    fn = obtain_fn()

    ctx.computing.add(address)
    try:
        try:
            value = fn(ctx)
        except XlErrorException as exc:
            ctx.cache[address] = exc.code
            raise
        if value is None and not (
            preserve_structural_blank and getattr(fn, "__structural_blank__", False)
        ):
            value = 0
        ctx.cache[address] = value
        return _raise_if_error_value(value)
    finally:
        ctx.computing.discard(address)

def xl_cell(ctx: EvalContextBase, address: str) -> CellValue:
    """Evaluate a single cell address under the given context (slim scaffold)."""

    def obtain_fn() -> Callable[[EvalContextBase], CellValue]:
        fn = ctx.resolver(address)
        if fn is None:
            raise KeyError(f"Cell {address} not found in graph")
        return cast(Callable[[EvalContextBase], CellValue], fn)

    return _evaluate_address(ctx, address, obtain_fn, preserve_structural_blank=True)

def _ctx_range(ctx: EvalContext, sheet: str, r1: int, c1: int, r2: int, c2: int) -> Range:
    # Leave the resolver unannotated: embed strips `excel_grapher.core` imports, so
    # aliases like `CellValue as CoreCellValue` never appear in generated runtime.py.
    def resolve(address: str):
        return xl_cell(ctx, address)

    return Range(sheet, r1, c1, r2, c2, resolve)

def xl_columns(ref: CellValue) -> int:
    """Return the column count of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(columns_count(ref))

def xl_compare(op: str, left: CellValue, right: CellValue) -> bool:
    """Compare two scalar operands with Excel ordering rules."""
    result = compare_scalars(op, as_scalar(left), as_scalar(right))
    if isinstance(result, XlError):
        raise _raise_error(result)
    return result

def xl_countif(range_values: CellValue, criteria: CellValue) -> int:
    """Count cells matching criteria, raising on Excel errors."""
    return raise_if_sentinel_int(countif_cells(range_values, criteria))

def xl_eval(
    ctx: EvalContextBase,
    address: str,
    fn: Callable[[EvalContextBase], CellValue],
) -> CellValue:
    """Evaluate a known formula implementation under the given context (slim scaffold)."""
    return _evaluate_address(ctx, address, lambda: fn, preserve_structural_blank=False)

def xl_hlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    row_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return _raise_if_error(hlookup_cells(lookup_value, table_array, row_index_num, range_lookup))

def xl_iferror(
    value_fn: Callable[[], CellValue], fallback_fn: Callable[[], CellValue]
) -> CellValue:
    """Excel IFERROR over lazily-evaluated value and fallback thunks."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return fallback_fn()
    if isinstance(value, XlError):
        return fallback_fn()
    return value

def xl_ifna(value_fn: Callable[[], CellValue], fallback_fn: Callable[[], CellValue]) -> CellValue:
    """Excel IFNA: catch `#N/A` only; other Excel errors propagate."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException as exc:
        if exc.code == XlError.NA:
            return fallback_fn()
        raise
    if value == XlError.NA:
        return fallback_fn()
    return value

def xl_index(array: CellValue, row_num: CellValue = None, col_num: CellValue = None) -> CellValue:
    return _raise_if_error(index_cells(array, row_num, col_num))

def xl_index_ref(
    ref: ExcelRange | tuple[str, int, int] | tuple[str, int, int, int, int],
    row_num: CellValue | None,
    col_num: CellValue | None,
) -> tuple[str, int, int] | tuple[str, int, int, int, int]:
    """Return INDEX reference metadata, raising on Excel reference errors."""
    out = index_excel_range(
        _range_from_ref_info(ref),
        _as_addressing_scalar(row_num),
        _as_addressing_scalar(col_num),
    )
    if isinstance(out, XlError):
        raise XlErrorException(out)
    if out.start_row == out.end_row and out.start_col == out.end_col:
        return (out.sheet, out.start_row, out.start_col)
    return (out.sheet, out.start_row, out.start_col, out.end_row, out.end_col)

def xl_int(value: CellValue) -> int:
    """Coerce a scalar cell value to an integer, raising on Excel errors."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    integer = to_int(scalar)
    if isinstance(integer, XlError):
        raise _raise_error(integer)
    return integer

def xl_is_array(value: object) -> bool:
    """Return whether *value* is a range or nested-list array operand."""
    return isinstance(value, (Range, ExcelRange, list, tuple))

def xl_isblank(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISBLANK: IS functions do not propagate errors."""
    try:
        value = value_fn()
    except XlErrorException:
        return False
    if isinstance(value, Range):
        if value.shape != (1, 1):
            return False
        value = value.value_at(1, 1)
    return value is None

def xl_iserror(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISERROR: True when evaluating the argument produces any Excel error."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException:
        return True
    return isinstance(value, XlError)

def xl_isna(value_fn: Callable[[], CellValue]) -> bool:
    """Excel ISNA: True when evaluating the argument produces `#N/A`."""
    try:
        value = _resolve_scalar(value_fn)
    except XlErrorException as exc:
        return exc.code == XlError.NA
    return value == XlError.NA

def xl_isnumber(value: CellValue) -> bool:
    return not isinstance(value, bool) and isinstance(value, (int, float))

def xl_large(array: CellValue, k: CellValue) -> float:
    """Return the k-th largest value, raising on Excel errors."""
    return raise_if_sentinel_float(large_kth(array, k))

def xl_left(text: CellValue, num_chars: CellValue = 1) -> str:
    """Return the leftmost characters of text, raising on Excel errors."""
    return raise_if_sentinel_str(left_chars(text, num_chars))

def xl_lookup(
    lookup_value: CellValue,
    lookup_vector_or_array: CellValue,
    result_vector: CellValue = None,
) -> CellValue:
    return _raise_if_error(lookup_cells(lookup_value, lookup_vector_or_array, result_vector))

def xl_map_compare(op: str, left: CellValue, right: CellValue) -> CellValue:
    """Element-wise comparison over scalar or broadcast array operands."""
    pair = _broadcast_pair(left, right)
    if pair is None:
        return xl_compare(op, left, right)

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            cell = compare_scalars(
                op, _cell_or_raise(arr_left, row0, col0), _cell_or_raise(arr_right, row0, col0)
            )
            if isinstance(cell, XlError):
                raise _raise_error(cell)
            out_row.append(cell)
        result.append(out_row)
    return result

def xl_map_concat(left: CellValue, right: CellValue) -> CellValue:
    """Element-wise string concatenation over scalar or broadcast array operands."""
    pair = _broadcast_pair(left, right)
    if pair is None:
        return concat_scalars(_scalar_or_raise(left), _scalar_or_raise(right))

    arr_left, arr_right = pair
    return [
        [
            concat_scalars(
                _cell_or_raise(arr_left, row0, col0), _cell_or_raise(arr_right, row0, col0)
            )
            for col0 in range(arr_left.ncols)
        ]
        for row0 in range(arr_left.nrows)
    ]

def xl_match(lookup_value: CellValue, lookup_array: CellValue, match_type: CellValue = 1) -> int:
    result = _raise_if_error(match_cells(lookup_value, lookup_array, match_type))
    return cast(int, result)

def xl_max(*args: CellValue) -> float:
    """Return the maximum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(max_cells(*args))

def xl_min(*args: CellValue) -> float:
    """Return the minimum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(min_cells(*args))

def xl_npv(rate: CellValue, *values: CellValue) -> float:
    """Return net present value, raising on Excel errors."""
    return raise_if_sentinel_float(npv_cells(rate, *values))

def xl_number(value: CellValue) -> float:
    """Coerce a scalar cell value to a number, raising on Excel errors."""
    scalar = as_scalar(value)
    if isinstance(scalar, XlError):
        raise _raise_error(scalar)
    number = to_number(scalar)
    if isinstance(number, XlError):
        raise _raise_error(number)
    return number

def xl_numbervalue(
    text: CellValue,
    decimal_separator: CellValue = ".",
    group_separator: CellValue = ",",
) -> float:
    """Convert text to a number, raising on Excel errors."""
    return raise_if_sentinel_float(numbervalue_parse(text, decimal_separator, group_separator))

def xl_offset(
    ctx: EvalContext,
    ref_info: tuple[str, int, int] | tuple[str, int, int, int, int] | XlError,
    rows: CellValue,
    cols: CellValue,
    height: CellValue | None = None,
    width: CellValue | None = None,
) -> CellValue:
    rr = _number_or_raise(rows)
    cc = _number_or_raise(cols)

    if isinstance(ref_info, XlError):
        raise XlErrorException(ref_info)

    match ref_info:
        case (sheet, base_row, base_col):
            base_end_row, base_end_col = base_row, base_col
        case (sheet, base_row, base_col, base_end_row, base_end_col):
            pass
        case _:
            raise XlErrorException(XlError.VALUE)

    base_h = int(base_end_row - base_row + 1)
    base_w = int(base_end_col - base_col + 1)

    h = base_h if height is None else int(_number_or_raise(height))
    w = base_w if width is None else int(_number_or_raise(width))

    target_row = int(base_row + int(rr))
    target_col = int(base_col + int(cc))

    if target_row < 1 or target_col < 1:
        raise XlErrorException(XlError.REF)
    if h <= 0 or w <= 0:
        raise XlErrorException(XlError.VALUE)

    if h == 1 and w == 1:
        addr = _format_address(sheet, target_row, target_col)
        # Scalar OFFSET results are CellValue; multi-cell returns a lazy Range.
        return cast("CellValue", xl_cell(ctx, addr))

    return _ctx_range(ctx, sheet, target_row, target_col, target_row + h - 1, target_col + w - 1)

def xl_or(*args: CellValue) -> bool:
    """Return logical OR, raising on Excel errors."""
    return raise_if_sentinel_bool(logical_or(*args))

def xl_pow_numbers(left: float, right: float) -> float:
    """Apply Excel exponentiation to coerced numbers."""
    try:
        value = left**right
    except (ValueError, OverflowError):
        raise _raise_error(XlError.NUM) from None
    if isinstance(value, complex):
        raise _raise_error(XlError.NUM)
    return value

def _apply_arithmetic_or_raise(op: str, left: CellValue, right: CellValue) -> CellValue:
    ln = xl_number(left)
    rn = xl_number(right)
    if op == "^":
        return xl_pow_numbers(ln, rn)
    cell = apply_arithmetic(op, ln, rn)
    if isinstance(cell, XlError):
        raise _raise_error(cell)
    return cell

def xl_map_arithmetic(op: str, left: CellValue, right: CellValue) -> CellValue:
    """Element-wise arithmetic over scalar or broadcast array operands."""
    pair = _broadcast_pair(left, right)
    if pair is None:
        return _apply_arithmetic_or_raise(op, left, right)

    arr_left, arr_right = pair
    result: list[list[CellValue]] = []
    for row0 in range(arr_left.nrows):
        out_row: list[CellValue] = []
        for col0 in range(arr_left.ncols):
            out_row.append(
                _apply_arithmetic_or_raise(
                    op, _cell_or_raise(arr_left, row0, col0), _cell_or_raise(arr_right, row0, col0)
                )
            )
        result.append(out_row)
    return result

def xl_raise(code: XlError) -> NoReturn:
    """Raise an Excel error code from an expression position."""
    raise XlErrorException(code)

def xl_range(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range address into a lazy `Range` value."""
    parsed = _parse_range_address(address)
    if isinstance(parsed, XlError):
        raise XlErrorException(parsed)
    sheet, start_cell, end_cell = parsed
    try:
        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_cell)
        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_cell)
        start_col_idx = fastpyxl.utils.cell.column_index_from_string(start_col)
        end_col_idx = fastpyxl.utils.cell.column_index_from_string(end_col)
    except ValueError:
        raise XlErrorException(XlError.VALUE) from None

    if start_row > end_row:
        start_row, end_row = end_row, start_row
    if start_col_idx > end_col_idx:
        start_col_idx, end_col_idx = end_col_idx, start_col_idx

    return _ctx_range(ctx, sheet, start_row, start_col_idx, end_row, end_col_idx)

def xl_range_rows(ctx: EvalContext, address: str) -> CellValue:
    """Evaluate a sheet-qualified range eagerly into nested row lists.

    Public boundary handler for range targets: results returned from
    `compute_all` are materialized values, not lazy range views.
    """
    rng = xl_range(ctx, address)
    if isinstance(rng, Range):
        return rng.rows_raw()
    return rng

def xl_rank(number: CellValue, ref: CellValue, order: CellValue = 0) -> int:
    """Return the rank of a number in a list, raising on Excel errors."""
    return raise_if_sentinel_int(rank_number(number, ref, order))

def xl_round(number: CellValue, num_digits: CellValue) -> float:
    """Round a number, raising on Excel coercion errors."""
    return raise_if_sentinel_float(round_number(number, num_digits))

def xl_rounddown(number: CellValue, num_digits: CellValue) -> float:
    """Round a number down, raising on Excel coercion errors."""
    return raise_if_sentinel_float(rounddown_number(number, num_digits))

def xl_row(ref: CellValue) -> int:
    """Return the row number of a reference, raising on Excel errors."""
    return raise_if_sentinel_int(row_number(ref))

def xl_stdev(*args: CellValue) -> float:
    """Return sample standard deviation, raising on Excel errors."""
    return raise_if_sentinel_float(stdev_cells(*args))

def xl_sum(*args: CellValue) -> float:
    """Return the sum of numeric cells, raising on Excel errors."""
    return raise_if_sentinel_float(sum_cells(*args))

def xl_sumproduct(*args: CellValue) -> float:
    if len(args) == 0:
        return 0.0
    grids: list[Grid] = []
    for arg in args:
        grid = Grid.wrap(arg)
        if grid is None:
            scalar_grid = Grid.wrap([[arg]])
            assert scalar_grid is not None
            grid = scalar_grid
        grids.append(grid)
    shape = (grids[0].nrows, grids[0].ncols)
    for grid in grids[1:]:
        if (grid.nrows, grid.ncols) != shape:
            raise XlErrorException(XlError.VALUE)

    result = 0.0
    for index0 in range(grids[0].size):
        product = 1.0
        for grid in grids:
            number = to_number(grid.at_flat(index0))
            if isinstance(number, XlError):
                raise XlErrorException(number)
            product *= number
        result += product
    return result

def xl_vlookup(
    lookup_value: CellValue,
    table_array: CellValue,
    col_index_num: CellValue,
    range_lookup: CellValue = True,
) -> CellValue:
    return _raise_if_error(vlookup_cells(lookup_value, table_array, col_index_num, range_lookup))

def xlookup_cells(
    lookup_value: object,
    lookup_array: object,
    return_array: object,
    if_not_found: object = None,
    match_mode: object = 0,
    search_mode: object = 1,
) -> object:
    """Excel XLOOKUP (exact match; search forward or backward)."""
    mm = to_number(cast(CellValue, match_mode))
    if isinstance(mm, XlError):
        return mm
    sm = to_number(cast(CellValue, search_mode))
    if isinstance(sm, XlError):
        return sm

    mm_i = int(mm)
    sm_i = int(sm)

    if mm_i != 0:
        return XlError.VALUE
    if sm_i not in (1, -1):
        return XlError.VALUE

    keys = Grid.wrap(lookup_array)
    vals = Grid.wrap(return_array)
    if keys is None or vals is None:
        return XlError.VALUE
    if keys.size != vals.size:
        return XlError.VALUE

    idxs = range(keys.size) if sm_i == 1 else range(keys.size - 1, -1, -1)
    for i in idxs:
        if _values_match(lookup_value, keys.at_flat(i)):
            return vals.at_flat(i)

    return XlError.NA if if_not_found is None else if_not_found

def xl_xlookup(
    lookup_value: CellValue,
    lookup_array: CellValue,
    return_array: CellValue,
    if_not_found: CellValue = None,
    match_mode: CellValue = 0,
    search_mode: CellValue = 1,
) -> CellValue:
    return _raise_if_error(
        xlookup_cells(
            lookup_value,
            lookup_array,
            return_array,
            if_not_found,
            match_mode,
            search_mode,
        )
    )
