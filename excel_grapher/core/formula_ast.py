from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from enum import StrEnum
from typing import TypeAlias

from fastpyxl.utils.cell import column_index_from_string, coordinate_from_string, get_column_letter

from excel_grapher.core.address_keys import (
    CellKey,
    format_cell_key,
    format_range_key,
    parse_address,
    quote_sheet_if_needed,
)

from .excel_function_names import normalize_excel_function_name
from .types import XlError


class FormulaParseError(Exception):
    """Raised when a formula cannot be parsed into an AST."""

    def __init__(self, formula: str, message: str) -> None:
        super().__init__(f"Parse error: {message}. Formula: {formula!r}")
        self.formula = formula
        self.message = message


@dataclass(frozen=True, slots=True)
class NumberNode:
    value: float


@dataclass(frozen=True, slots=True)
class StringNode:
    value: str


@dataclass(frozen=True, slots=True)
class BoolNode:
    value: bool


@dataclass(frozen=True, slots=True)
class ErrorNode:
    error: XlError


@dataclass(frozen=True, slots=True)
class AbsoluteAxis:
    """1-based absolute row number or column index."""

    index: int


@dataclass(frozen=True, slots=True)
class RelativeAxis:
    """Offset from the host cell along one axis."""

    offset: int


AxisRef: TypeAlias = AbsoluteAxis | RelativeAxis


class FormulaStyle(StrEnum):
    """How `render_formula` spells cell and range references.

    `A1_ABSOLUTE` is the `normalized_formula` dialect: sheet-qualified A1 with
    `$` stripped. `A1_EXCEL` keeps `$` on absolute axes and omits the host sheet
    prefix. `R1C1` uses `R1C1` / `R[-1]C[2]` tokens.
    """

    A1_ABSOLUTE = "a1_absolute"
    A1_EXCEL = "a1_excel"
    R1C1 = "r1c1"


@dataclass(frozen=True, slots=True)
class CellRef:
    """Sheet-qualified cell with per-axis relative/absolute intent."""

    sheet: str
    col: AxisRef
    row: AxisRef


def cell_ref_from_a1(address: str) -> CellRef:
    """Build a fully-absolute `CellRef` from sheet-qualified A1 text."""
    sheet, coord = parse_address(address)
    col_letter, row = coordinate_from_string(coord.replace("$", ""))
    return CellRef(
        sheet=sheet,
        col=AbsoluteAxis(int(column_index_from_string(col_letter))),
        row=AbsoluteAxis(int(row)),
    )


@dataclass(frozen=True, slots=True)
class CellRefNode:
    """AST leaf for a sheet-qualified cell reference."""

    ref: CellRef

    def __init__(self, ref: CellRef | str | None = None, /, *, address: str | None = None) -> None:
        if address is not None:
            if ref is not None:
                raise TypeError("pass ref or address, not both")
            object.__setattr__(self, "ref", cell_ref_from_a1(address))
        elif isinstance(ref, str):
            object.__setattr__(self, "ref", cell_ref_from_a1(ref))
        elif isinstance(ref, CellRef):
            object.__setattr__(self, "ref", ref)
        else:
            raise TypeError("CellRefNode requires a CellRef or A1 address")

    @property
    def address(self) -> str:
        """Canonical sheet-qualified A1 when both axes are absolute."""
        return resolve_cell_ref(self.ref, None)


@dataclass(frozen=True, slots=True)
class RangeNode:
    """AST leaf for an A1 range with per-endpoint axis intent."""

    start_ref: CellRef
    end_ref: CellRef

    def __init__(
        self,
        start: str | CellRef | None = None,
        end: str | CellRef | None = None,
        *,
        start_ref: CellRef | None = None,
        end_ref: CellRef | None = None,
    ) -> None:
        raw_start = start_ref if start_ref is not None else start
        raw_end = end_ref if end_ref is not None else end
        if raw_start is None or raw_end is None:
            raise TypeError("RangeNode requires start and end")
        object.__setattr__(
            self,
            "start_ref",
            raw_start if isinstance(raw_start, CellRef) else cell_ref_from_a1(raw_start),
        )
        object.__setattr__(
            self,
            "end_ref",
            raw_end if isinstance(raw_end, CellRef) else cell_ref_from_a1(raw_end),
        )

    @property
    def start(self) -> str:
        """Canonical start A1 when both start axes are absolute."""
        return resolve_cell_ref(self.start_ref, None)

    @property
    def end(self) -> str:
        """Canonical end A1 when both end axes are absolute."""
        return resolve_cell_ref(self.end_ref, None)


@dataclass(frozen=True, slots=True)
class WholeColumnNode:
    sheet: str
    col: AxisRef

    def __init__(
        self,
        sheet: str,
        column: str | AxisRef | None = None,
        *,
        col: AxisRef | None = None,
    ) -> None:
        object.__setattr__(self, "sheet", sheet)
        axis = col if col is not None else column
        if axis is None:
            raise TypeError("WholeColumnNode requires a column")
        if isinstance(axis, str):
            axis = AbsoluteAxis(int(column_index_from_string(axis.upper())))
        object.__setattr__(self, "col", axis)

    @property
    def column(self) -> str:
        """Column letter when this whole-column ref is absolute."""
        _sheet, letter = resolve_whole_column_ref(self, None)
        return letter


@dataclass(frozen=True, slots=True)
class WholeRowNode:
    sheet: str
    row: AxisRef

    def __init__(self, sheet: str, row: int | AxisRef) -> None:
        object.__setattr__(self, "sheet", sheet)
        if isinstance(row, int):
            object.__setattr__(self, "row", AbsoluteAxis(row))
        else:
            object.__setattr__(self, "row", row)


def _resolve_axis(axis: AxisRef, base: int | None) -> int:
    if isinstance(axis, AbsoluteAxis):
        return axis.index
    if base is None:
        raise ValueError("relative axis requires an anchor cell")
    return base + axis.offset


def _coerce_anchor_key(anchor: CellKey | str | None) -> CellKey | None:
    if anchor is None:
        return None
    if isinstance(anchor, CellKey):
        return anchor
    return CellKey(str(anchor))


def resolve_cell_ref(ref: CellRef | CellRefNode, anchor: CellKey | str | None) -> str:
    """Resolve `ref` to canonical sheet-qualified A1.

    Absolute axes ignore `anchor`. Relative axes add their offset to `anchor`.

    Raises:
        ValueError: If a relative axis is present and `anchor` is missing, or
            the resolved row/column is less than 1.
    """
    cell = ref.ref if isinstance(ref, CellRefNode) else ref
    anchor_key = _coerce_anchor_key(anchor)
    col_base = None if anchor_key is None else int(column_index_from_string(anchor_key.column))
    row_base = None if anchor_key is None else int(anchor_key.row)
    col_index = _resolve_axis(cell.col, col_base)
    row_index = _resolve_axis(cell.row, row_base)
    if col_index < 1 or row_index < 1:
        raise ValueError(f"resolved address out of range: col={col_index} row={row_index}")
    return format_cell_key(cell.sheet, get_column_letter(col_index), row_index)


def resolve_whole_column_ref(
    node: WholeColumnNode, anchor: CellKey | str | None
) -> tuple[str, str]:
    """Resolve a whole-column leaf to `(sheet, column_letter)`."""
    anchor_key = _coerce_anchor_key(anchor)
    col_base = None if anchor_key is None else int(column_index_from_string(anchor_key.column))
    col_index = _resolve_axis(node.col, col_base)
    if col_index < 1:
        raise ValueError(f"resolved column out of range: {col_index}")
    return node.sheet, get_column_letter(col_index)


def resolve_whole_row_ref(node: WholeRowNode, anchor: CellKey | str | None) -> tuple[str, int]:
    """Resolve a whole-row leaf to `(sheet, row_number)`."""
    anchor_key = _coerce_anchor_key(anchor)
    row_base = None if anchor_key is None else int(anchor_key.row)
    row_index = _resolve_axis(node.row, row_base)
    if row_index < 1:
        raise ValueError(f"resolved row out of range: {row_index}")
    return node.sheet, row_index


@dataclass(frozen=True, slots=True)
class FunctionCallNode:
    """Function invocation. `args` is frozen so formula trees are hashable."""

    name: str
    args: tuple[AstNode, ...]

    def __init__(self, name: str, args: Sequence[AstNode]) -> None:
        object.__setattr__(self, "name", name)
        object.__setattr__(self, "args", tuple(args))


@dataclass(frozen=True, slots=True)
class BinaryOpNode:
    op: str
    left: AstNode
    right: AstNode


@dataclass(frozen=True, slots=True)
class UnaryOpNode:
    op: str
    operand: AstNode


@dataclass(frozen=True, slots=True)
class EmptyArgNode:
    """Represents an omitted argument in a function call (e.g., INDEX(A1:B2,,1))."""

    pass


AstNode: TypeAlias = (
    NumberNode
    | StringNode
    | BoolNode
    | ErrorNode
    | CellRefNode
    | RangeNode
    | WholeColumnNode
    | WholeRowNode
    | FunctionCallNode
    | BinaryOpNode
    | UnaryOpNode
    | EmptyArgNode
)


def intern_formula_ast(tree: AstNode, intern: dict[AstNode, AstNode]) -> AstNode:
    """Return the canonical interned instance of `tree`.

    The intern map is keyed by the frozen tree itself. Do not intern a JSON
    encoding of `tree`.
    """
    return intern.setdefault(tree, tree)


def iter_resolved_cell_keys(node: AstNode, anchor: CellKey | str) -> Iterator[str]:
    """Yield canonical cell keys referenced by `node` against `anchor`.

    Range endpoints are yielded (not expanded). Whole-column/row leaves are
    skipped; callers that need those bounds should resolve them separately.
    """
    match node:
        case CellRefNode(ref):
            yield resolve_cell_ref(ref, anchor)
        case RangeNode(start_ref, end_ref):
            yield resolve_cell_ref(start_ref, anchor)
            yield resolve_cell_ref(end_ref, anchor)
        case FunctionCallNode(_, args):
            for arg in args:
                yield from iter_resolved_cell_keys(arg, anchor)
        case BinaryOpNode(_, left, right):
            yield from iter_resolved_cell_keys(left, anchor)
            yield from iter_resolved_cell_keys(right, anchor)
        case UnaryOpNode(_, operand):
            yield from iter_resolved_cell_keys(operand, anchor)
        case _:
            return


def bind_axes(node: AstNode, anchor: CellKey | str | None) -> AstNode:
    """Return a copy of `node` with every axis resolved to `AbsoluteAxis`."""
    match node:
        case CellRefNode(ref):
            return CellRefNode(resolve_cell_ref(ref, anchor))
        case RangeNode(start_ref, end_ref):
            return RangeNode(
                resolve_cell_ref(start_ref, anchor),
                resolve_cell_ref(end_ref, anchor),
            )
        case WholeColumnNode():
            sheet, letter = resolve_whole_column_ref(node, anchor)
            return WholeColumnNode(sheet=sheet, column=letter)
        case WholeRowNode():
            sheet, row = resolve_whole_row_ref(node, anchor)
            return WholeRowNode(sheet=sheet, row=row)
        case FunctionCallNode(name, args):
            return FunctionCallNode(name, [bind_axes(arg, anchor) for arg in args])
        case BinaryOpNode(op, left, right):
            return BinaryOpNode(op, bind_axes(left, anchor), bind_axes(right, anchor))
        case UnaryOpNode(op, operand):
            return UnaryOpNode(op, bind_axes(operand, anchor))
        case _:
            return node


def _unparse_number(value: float) -> str:
    as_float = float(value)
    if as_float.is_integer() and abs(as_float) < 1e15:
        return str(int(as_float))
    return format(as_float, ".15g")


def _unparse_string(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def _host_sheet(anchor: CellKey | str | None) -> str | None:
    key = _coerce_anchor_key(anchor)
    return None if key is None else key.sheet


def _sheet_prefix(sheet: str, *, style: FormulaStyle, host_sheet: str | None) -> str:
    if style is FormulaStyle.A1_ABSOLUTE or host_sheet is None or sheet != host_sheet:
        return f"{quote_sheet_if_needed(sheet)}!"
    return ""


def _r1c1_axis(axis: AxisRef, *, is_row: bool) -> str:
    token = "R" if is_row else "C"
    if isinstance(axis, AbsoluteAxis):
        return f"{token}{axis.index}"
    if axis.offset == 0:
        return token
    return f"{token}[{axis.offset}]"


def _r1c1_cell(ref: CellRef) -> str:
    return _r1c1_axis(ref.row, is_row=True) + _r1c1_axis(ref.col, is_row=False)


def _a1_excel_coord(ref: CellRef, anchor: CellKey | str | None) -> str:
    key = _coerce_anchor_key(anchor)
    col_base = None if key is None else int(column_index_from_string(key.column))
    row_base = None if key is None else int(key.row)
    col_index = _resolve_axis(ref.col, col_base)
    row_index = _resolve_axis(ref.row, row_base)
    if col_index < 1 or row_index < 1:
        raise ValueError(f"resolved address out of range: col={col_index} row={row_index}")
    letter = get_column_letter(col_index)
    col_s = f"${letter}" if isinstance(ref.col, AbsoluteAxis) else letter
    row_s = f"${row_index}" if isinstance(ref.row, AbsoluteAxis) else str(row_index)
    return f"{col_s}{row_s}"


def _qualify_atom(sheet: str, body: str, *, style: FormulaStyle, host_sheet: str | None) -> str:
    return f"{_sheet_prefix(sheet, style=style, host_sheet=host_sheet)}{body}"


def _unparse_atom_ref(
    node: CellRefNode | RangeNode | WholeColumnNode | WholeRowNode,
    anchor: CellKey | str | None,
    style: FormulaStyle,
) -> str:
    host_sheet = _host_sheet(anchor)
    match node:
        case CellRefNode(ref):
            if style is FormulaStyle.R1C1:
                return _qualify_atom(ref.sheet, _r1c1_cell(ref), style=style, host_sheet=host_sheet)
            if style is FormulaStyle.A1_EXCEL:
                return _qualify_atom(
                    ref.sheet, _a1_excel_coord(ref, anchor), style=style, host_sheet=host_sheet
                )
            return resolve_cell_ref(ref, anchor)
        case RangeNode(start_ref, end_ref):
            if style is FormulaStyle.R1C1:
                start_body = _r1c1_cell(start_ref)
                end_body = _r1c1_cell(end_ref)
            elif style is FormulaStyle.A1_EXCEL:
                start_body = _a1_excel_coord(start_ref, anchor)
                end_body = _a1_excel_coord(end_ref, anchor)
            else:
                start = resolve_cell_ref(start_ref, anchor)
                end = resolve_cell_ref(end_ref, anchor)
                start_sheet, start_coord = parse_address(start)
                end_sheet, end_coord = parse_address(end)
                if start_sheet == end_sheet:
                    return format_range_key(start_sheet, start_coord, end_coord)
                return f"{start}:{end}"
            if start_ref.sheet == end_ref.sheet:
                return (
                    f"{_qualify_atom(start_ref.sheet, start_body, style=style, host_sheet=host_sheet)}"
                    f":{end_body}"
                )
            return (
                f"{_qualify_atom(start_ref.sheet, start_body, style=style, host_sheet=host_sheet)}:"
                f"{_qualify_atom(end_ref.sheet, end_body, style=style, host_sheet=host_sheet)}"
            )
        case WholeColumnNode():
            if style is FormulaStyle.R1C1:
                token = _r1c1_axis(node.col, is_row=False)
                body = f"{token}:{token}"
                return _qualify_atom(node.sheet, body, style=style, host_sheet=host_sheet)
            sheet, letter = resolve_whole_column_ref(node, anchor)
            if style is FormulaStyle.A1_EXCEL:
                marked = f"${letter}" if isinstance(node.col, AbsoluteAxis) else letter
                return _qualify_atom(
                    sheet, f"{marked}:{marked}", style=style, host_sheet=host_sheet
                )
            return f"{quote_sheet_if_needed(sheet)}!{letter}:{letter}"
        case WholeRowNode():
            if style is FormulaStyle.R1C1:
                token = _r1c1_axis(node.row, is_row=True)
                body = f"{token}:{token}"
                return _qualify_atom(node.sheet, body, style=style, host_sheet=host_sheet)
            sheet, row = resolve_whole_row_ref(node, anchor)
            if style is FormulaStyle.A1_EXCEL:
                marked = f"${row}" if isinstance(node.row, AbsoluteAxis) else str(row)
                return _qualify_atom(
                    sheet, f"{marked}:{marked}", style=style, host_sheet=host_sheet
                )
            return f"{quote_sheet_if_needed(sheet)}!{row}:{row}"


def _unparse_expr(
    node: AstNode,
    *,
    anchor: CellKey | str | None,
    style: FormulaStyle,
    parent_prec: int,
    is_right: bool,
) -> str:
    match node:
        case NumberNode(value):
            return _unparse_number(value)
        case StringNode(value):
            return _unparse_string(value)
        case BoolNode(value):
            return "TRUE" if value else "FALSE"
        case ErrorNode(error):
            return error.value
        case EmptyArgNode():
            return ""
        case CellRefNode() | RangeNode() | WholeColumnNode() | WholeRowNode():
            return _unparse_atom_ref(node, anchor, style)
        case FunctionCallNode(name, args):
            inner = ",".join(
                _unparse_expr(arg, anchor=anchor, style=style, parent_prec=0, is_right=False)
                for arg in args
            )
            return f"{name}({inner})"
        case UnaryOpNode("%", operand):
            body = _unparse_expr(operand, anchor=anchor, style=style, parent_prec=6, is_right=False)
            return f"{body}%"
        case UnaryOpNode(op, operand):
            body = _unparse_expr(operand, anchor=anchor, style=style, parent_prec=6, is_right=False)
            return f"{op}{body}"
        case BinaryOpNode(op, left, right):
            prec = _PRECEDENCE[op]
            left_s = _unparse_expr(
                left, anchor=anchor, style=style, parent_prec=prec, is_right=False
            )
            right_s = _unparse_expr(
                right, anchor=anchor, style=style, parent_prec=prec, is_right=True
            )
            text = f"{left_s}{op}{right_s}"
            need_parens = prec < parent_prec or (
                is_right and op not in _RIGHT_ASSOC and prec == parent_prec
            )
            return f"({text})" if need_parens else text
    raise TypeError(f"unsupported AST node: {type(node).__name__}")


def render_formula(
    ast: AstNode,
    *,
    anchor: CellKey | str | None = None,
    style: FormulaStyle | str = FormulaStyle.A1_ABSOLUTE,
    coerce_relative_refs: bool = False,
) -> str:
    """Render `ast` as formula text (leading `=`).

    This is the formula-text dialect of record. It is not a byte-identical
    copy of Excel's raw formula or of regex `normalize_excel_formula`. The
    tree is canonicalized on parse/render:

    - unary `+` is dropped (parse)
    - redundant parentheses are omitted; only precedence-required ones remain
    - number spelling is canonical (`1.0` -> `1`, `1e2` -> `100`)
    - whitespace is compact (`=SUM( A1 )` renders as `=SUM(Sheet1!A1)`)

    Bare refs are sheet-qualified when parsed with an `anchor`. `style`
    selects reference spelling. Relative axes resolve against `anchor`. When
    `coerce_relative_refs` is True, every `RelativeAxis` is bound to an
    `AbsoluteAxis` first (preparing Excel write-back that wants fully
    absolute addresses). Same-sheet prefixes are omitted for `A1_EXCEL` and
    `R1C1` when `anchor` is the host cell.

    Args:
        ast: Formula tree to stringify.
        anchor: Host cell for relative axes. Required when `ast` has any
            `RelativeAxis`, or when `coerce_relative_refs` is True and relative
            axes are present.
        style: `A1_ABSOLUTE`, `A1_EXCEL`, or `R1C1` (or the matching string).
        coerce_relative_refs: If True, bind relative axes to absolute indexes
            before spelling them.

    Returns:
        Formula text beginning with `=`.

    Raises:
        ValueError: If a relative axis is present and `anchor` is missing, or
            a resolved row/column is less than 1.
    """
    resolved_style = FormulaStyle(style)
    tree = bind_axes(ast, anchor) if coerce_relative_refs else ast
    return "=" + _unparse_expr(
        tree, anchor=anchor, style=resolved_style, parent_prec=0, is_right=False
    )


def unparse_normalized_formula(
    node: AstNode,
    *,
    anchor: CellKey | str | None = None,
) -> str:
    """Render `node` as absolute A1 formula text (leading `=`).

    Equivalent to `render_formula` with `style=A1_ABSOLUTE`. Relative axes
    resolve against `anchor`. Same-sheet ranges use a single sheet prefix.

    Args:
        node: Formula AST to render.
        anchor: Host cell for relative axes. Required when `node` has any
            `RelativeAxis`.

    Returns:
        Sheet-qualified absolute A1 formula beginning with `=`.
    """
    return render_formula(node, anchor=anchor, style=FormulaStyle.A1_ABSOLUTE)


def _retarget_cell_ref(ref: CellRef, new_key: str, anchor: CellKey | str) -> CellRef:
    """Point `ref` at `new_key` while keeping each axis's relative/absolute kind."""
    sheet, coord = parse_address(new_key)
    col_letter, row = coordinate_from_string(coord.replace("$", ""))
    new_col = int(column_index_from_string(col_letter))
    new_row = int(row)
    anchor_key = _coerce_anchor_key(anchor)
    if anchor_key is None:
        raise ValueError("retargeting a relative axis requires an anchor cell")
    col_base = int(column_index_from_string(anchor_key.column))
    row_base = int(anchor_key.row)

    def toward(axis: AxisRef, new_index: int, base: int) -> AxisRef:
        if isinstance(axis, AbsoluteAxis):
            return AbsoluteAxis(new_index)
        return RelativeAxis(new_index - base)

    return CellRef(
        sheet=sheet,
        col=toward(ref.col, new_col, col_base),
        row=toward(ref.row, new_row, row_base),
    )


def replace_resolved_cell_ref(
    node: AstNode,
    *,
    old_key: str,
    new_key: str,
    anchor: CellKey | str,
    replacement: AstNode | None = None,
) -> AstNode:
    """Replace `CellRefNode` leaves that resolve to `old_key`.

    Range and whole-column/row leaves are left unchanged. When `replacement` is
    omitted, matching leaves keep each axis's relative/absolute kind and retarget
    to `new_key` against `anchor`.

    Args:
        node: Formula AST to rewrite.
        old_key: Canonical sheet-qualified address to match after resolution.
        new_key: Replacement address when `replacement` is omitted.
        anchor: Host cell used to resolve relative axes and retarget offsets.
        replacement: Optional subtree to splice in place of each match.

    Returns:
        A new tree when any leaf changed; otherwise `node`.
    """

    def walk(cur: AstNode) -> AstNode:
        match cur:
            case CellRefNode(ref):
                if resolve_cell_ref(ref, anchor) == old_key:
                    if replacement is not None:
                        return replacement
                    return CellRefNode(_retarget_cell_ref(ref, new_key, anchor))
                return cur
            case FunctionCallNode(name, args):
                new_args = [walk(arg) for arg in args]
                if new_args == args:
                    return cur
                return FunctionCallNode(name, new_args)
            case BinaryOpNode(op, left, right):
                new_left = walk(left)
                new_right = walk(right)
                if new_left is left and new_right is right:
                    return cur
                return BinaryOpNode(op, new_left, new_right)
            case UnaryOpNode(op, operand):
                new_operand = walk(operand)
                if new_operand is operand:
                    return cur
                return UnaryOpNode(op, new_operand)
            case _:
                return cur

    return walk(node)


class _Scanner:
    def __init__(
        self,
        text: str,
        *,
        anchor: CellKey | None = None,
        preserve_axes: bool = False,
    ) -> None:
        self.text = text
        self.i = 0
        self.anchor = anchor
        self.preserve_axes = preserve_axes

    def peek(self) -> str | None:
        if self.i >= len(self.text):
            return None
        return self.text[self.i]

    def consume(self) -> str | None:
        ch = self.peek()
        if ch is None:
            return None
        self.i += 1
        return ch

    def skip_ws(self) -> None:
        while (c := self.peek()) is not None and c.isspace():
            self.i += 1

    def take_while(self, pred) -> str:
        start = self.i
        while (c := self.peek()) is not None and pred(c):
            self.i += 1
        return self.text[start : self.i]

    def eof(self) -> bool:
        return self.peek() is None

    @property
    def default_sheet(self) -> str | None:
        return None if self.anchor is None else self.anchor.sheet


# Operator precedence (higher = binds tighter)
# Excel precedence: comparison < concat < add/sub < mul/div < exponent < unary
_PRECEDENCE: dict[str, int] = {
    "=": 1,
    "<": 1,
    ">": 1,
    "<=": 1,
    ">=": 1,
    "<>": 1,
    "&": 2,
    "+": 3,
    "-": 3,
    "*": 4,
    "/": 4,
    "^": 5,
}

# Right-associative operators
_RIGHT_ASSOC: set[str] = {"^"}


def parse(
    formula: str,
    *,
    anchor: CellKey | str | None = None,
    preserve_axes: bool = False,
) -> AstNode:
    """Parse a formula into an AST.

    By default every cell/range axis is stored as `AbsoluteAxis` (the historical
    `parse(normalized_formula)` contract). Pass `preserve_axes=True` with
    `anchor` to keep `$` vs bare A1 as absolute vs relative offsets.

    Excel-like whitespace around operators, commas, and call arguments is
    accepted. Scientific literals (`1e2`, `1E+2`, `1.5e-1`) become `NumberNode`.
    Incomplete exponents (`1e`, `1E+`) raise `FormulaParseError` (fail-soft via
    `parse_optional` / `parse_preserving_axes_optional`). Unary `+` is dropped.

    Args:
        formula: Excel formula text, with or without a leading `=`.
        anchor: Host cell used when `preserve_axes` is True, and as the default
            sheet for unqualified refs.
        preserve_axes: If True, missing `$` becomes a `RelativeAxis` offset from
            `anchor`. Requires `anchor`.

    Raises:
        FormulaParseError: If `formula` is not a supported expression.
        ValueError: If `preserve_axes` is True and `anchor` is missing.
    """
    raw = formula.strip()
    if raw.startswith("="):
        raw = raw[1:].strip()

    coerced = _coerce_anchor_key(anchor)
    if preserve_axes and coerced is None:
        raise ValueError("preserve_axes requires an anchor cell")
    s = _Scanner(raw, anchor=coerced, preserve_axes=preserve_axes)
    node = _parse_expression(s, formula, min_prec=0)
    s.skip_ws()
    if not s.eof():
        raise FormulaParseError(formula, f"Unexpected trailing input at {s.i}")
    return node


def parse_optional(formula: str | None) -> AstNode | None:
    """Parse `formula`, or return None when it is missing, blank, or unparseable.

    Catches `FormulaParseError` from `parse`. Other exceptions still propagate.
    """
    if formula is None:
        return None
    stripped = formula.strip()
    if not stripped:
        return None
    try:
        return parse(stripped)
    except FormulaParseError:
        return None


def parse_preserving_axes(
    formula: str,
    *,
    anchor: CellKey | str,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> AstNode:
    """Parse `formula` from raw workbook text, preserving `$` axis intent.

    Bare A1 refs resolve against `anchor`. Defined names expand to absolute
    sheet-qualified A1 before parsing. Whitespace and scientific literals follow
    the same acceptance rules as `parse`.
    """
    from excel_grapher.core.formula_normalization import expand_defined_names

    expanded = expand_defined_names(
        formula,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )
    return parse(expanded, anchor=anchor, preserve_axes=True)


def parse_preserving_axes_optional(
    formula: str | None,
    *,
    anchor: CellKey | str,
    named_ranges: dict[str, tuple[str, str]] | None = None,
    named_range_ranges: dict[str, tuple[str, str, str]] | None = None,
) -> AstNode | None:
    """Like `parse_preserving_axes`, returning None on missing/blank/unparseable input."""
    if formula is None:
        return None
    stripped = formula.strip()
    if not stripped:
        return None
    try:
        return parse_preserving_axes(
            stripped,
            anchor=anchor,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
    except FormulaParseError:
        return None


def _parse_expression(s: _Scanner, original: str, min_prec: int) -> AstNode:
    """Pratt parser / precedence climbing for expressions."""
    left = _parse_unary(s, original)

    while True:
        s.skip_ws()
        op = _peek_operator(s)
        if op is None:
            break
        prec = _PRECEDENCE.get(op)
        if prec is None or prec < min_prec:
            break

        # Consume the operator
        for _ in range(len(op)):
            s.consume()

        # Right associativity: use same precedence; left associativity: use prec + 1
        next_min = prec if op in _RIGHT_ASSOC else prec + 1
        right = _parse_expression(s, original, next_min)
        left = BinaryOpNode(op, left, right)

    return left


def _peek_operator(s: _Scanner) -> str | None:
    """Peek at the next operator (may be 1 or 2 chars)."""
    ch = s.peek()
    if ch is None:
        return None

    # Check for two-character operators first
    if s.i + 1 < len(s.text):
        two = s.text[s.i : s.i + 2]
        if two in ("<=", ">=", "<>"):
            return two

    # Single-character operators
    if ch in _PRECEDENCE:
        return ch

    return None


def _parse_unary(s: _Scanner, original: str) -> AstNode:
    """Parse unary operators (-, +) and atoms."""
    s.skip_ws()
    ch = s.peek()

    # Unary minus
    if ch == "-":
        s.consume()
        operand = _parse_unary(s, original)
        return UnaryOpNode("-", operand)

    # Unary plus (just ignore it)
    if ch == "+":
        s.consume()
        return _parse_unary(s, original)

    node = _parse_atom(s, original)

    # Postfix percent operator: 100% -> 1.0
    while True:
        s.skip_ws()
        if s.peek() != "%":
            break
        s.consume()
        node = UnaryOpNode("%", node)

    return node


def _col_axis(s: _Scanner, col_index: int, is_abs: bool) -> AxisRef:
    if is_abs or not s.preserve_axes or s.anchor is None:
        return AbsoluteAxis(col_index)
    base = int(column_index_from_string(s.anchor.column))
    return RelativeAxis(col_index - base)


def _row_axis(s: _Scanner, row_index: int, is_abs: bool) -> AxisRef:
    if is_abs or not s.preserve_axes or s.anchor is None:
        return AbsoluteAxis(row_index)
    return RelativeAxis(row_index - int(s.anchor.row))


def _parse_atom(s: _Scanner, original: str) -> AstNode:
    """Parse an atomic expression (literal, cell ref, function call, or parenthesized expr)."""
    s.skip_ws()
    ch = s.peek()
    if ch is None:
        raise FormulaParseError(original, "Empty formula")

    # Parenthesized expression
    if ch == "(":
        s.consume()
        node = _parse_expression(s, original, min_prec=0)
        s.skip_ws()
        if s.peek() != ")":
            raise FormulaParseError(original, "Expected ')' after parenthesized expression")
        s.consume()
        return node

    if ch == '"':
        return _parse_string(s, original)

    if ch == "#":
        return _parse_error(s, original)

    if ch == "'":
        return _parse_quoted_sheet_ref(s, original)

    if ch == "$":
        return _parse_local_ref(s, original)

    if ch.isdigit() or ch == ".":
        saved = s.i
        whole_row = _try_parse_local_whole_row(s, original)
        if whole_row is not None:
            return whole_row
        s.i = saved
        return _parse_number(s, original)

    if ch.isalpha() or ch in ("_",):
        ident = _parse_ident(s)
        upper = ident.upper()

        s.skip_ws()
        if s.peek() == "(":
            s.consume()  # '('
            args = _parse_args(s, original)
            return FunctionCallNode(name=normalize_excel_function_name(upper), args=args)

        if upper == "TRUE":
            return BoolNode(True)
        if upper == "FALSE":
            return BoolNode(False)

        if s.peek() == "!":
            s.consume()
            return _parse_ref_after_sheet_bang(s, original, sheet_qualifier=ident)

        local = _try_finish_local_ident_ref(s, original, ident)
        if local is not None:
            return local

        raise FormulaParseError(
            original, "Cell references must be sheet-qualified (e.g., Sheet1!A1)"
        )

    raise FormulaParseError(original, f"Unexpected character {ch!r} at {s.i}")


def _local_sheet(s: _Scanner, original: str) -> str:
    if s.default_sheet is None:
        raise FormulaParseError(
            original, "Cell references must be sheet-qualified (e.g., Sheet1!A1)"
        )
    return s.default_sheet


def _try_parse_local_whole_row(s: _Scanner, original: str) -> WholeRowNode | None:
    if s.default_sheet is None:
        return None
    s.skip_ws()
    abs_row = False
    if s.peek() == "$":
        s.consume()
        abs_row = True
    row_ch = s.peek()
    if row_ch is None or not row_ch.isdigit():
        return None
    row_str = s.take_while(lambda c: c.isdigit())
    s.skip_ws()
    if s.peek() != ":":
        return None
    s.consume()
    s.skip_ws()
    if s.peek() == "$":
        s.consume()
        abs_row = True
    row2 = s.take_while(lambda c: c.isdigit())
    if row2 != row_str:
        return None
    return WholeRowNode(sheet=s.default_sheet, row=_row_axis(s, int(row_str), abs_row))


def _parse_local_ref(s: _Scanner, original: str) -> AstNode:
    """Parse a bare `$`-prefixed local ref (`$A$1`, `$A:A`, `$1:$1`)."""
    sheet = _local_sheet(s, original)
    s.skip_ws()
    if s.peek() != "$":
        raise FormulaParseError(original, f"Expected '$' at {s.i}")
    saved = s.i
    s.consume()
    ch = s.peek()
    if ch is not None and ch.isdigit():
        s.i = saved
        whole_row = _try_parse_local_whole_row(s, original)
        if whole_row is None:
            raise FormulaParseError(original, f"Invalid cell coordinate at {s.i}")
        return whole_row
    s.i = saved
    return _parse_a1_or_whole_col_or_range(s, original, sheet)


def _try_finish_local_ident_ref(s: _Scanner, original: str, ident: str) -> AstNode | None:
    if s.default_sheet is None:
        return None
    sheet = s.default_sheet
    if ident.isalpha() and s.peek() == ":":
        s.i -= len(ident)
        return _parse_a1_or_whole_col_or_range(s, original, sheet)
    if ident.isalpha() and s.peek() == "$":
        s.i -= len(ident)
        return _parse_a1_or_whole_col_or_range(s, original, sheet)
    col_row = _split_a1_ident(ident)
    if col_row is not None:
        col_letters, row_str = col_row
        ref = CellRef(
            sheet=sheet,
            col=_col_axis(s, int(column_index_from_string(col_letters)), is_abs=False),
            row=_row_axis(s, int(row_str), is_abs=False),
        )
        s.skip_ws()
        if s.peek() == ":":
            s.consume()
            end = _parse_range_end_ref(s, original, default_sheet=sheet)
            return RangeNode(start_ref=ref, end_ref=end)
        return CellRefNode(ref)
    return None


def _split_a1_ident(ident: str) -> tuple[str, str] | None:
    i = 0
    while i < len(ident) and ident[i].isalpha():
        i += 1
    if i == 0 or i == len(ident):
        return None
    col, row = ident[:i], ident[i:]
    if not row.isdigit() or not (1 <= len(col) <= 3):
        return None
    return col.upper(), row


def _parse_a1_or_whole_col_or_range(s: _Scanner, original: str, sheet: str) -> AstNode:
    saved = s.i
    whole_col = _try_parse_whole_column(s, original, sheet)
    if whole_col is not None:
        return whole_col
    s.i = saved
    return _parse_cell_or_range(s, original, sheet)


def _try_parse_whole_column(s: _Scanner, original: str, sheet: str) -> WholeColumnNode | None:
    del original
    s.skip_ws()
    abs_col = False
    if s.peek() == "$":
        s.consume()
        abs_col = True
    col = s.take_while(lambda c: c.isalpha())
    if not col:
        return None
    s.skip_ws()
    if s.peek() != ":":
        return None
    s.consume()
    s.skip_ws()
    if s.peek() == "$":
        s.consume()
        abs_col = True
    col2 = s.take_while(lambda c: c.isalpha())
    if col2.upper() != col.upper():
        return None
    after_col = s.peek()
    if after_col is not None and after_col.isdigit():
        return None
    return WholeColumnNode(
        sheet=sheet,
        col=_col_axis(s, int(column_index_from_string(col.upper())), abs_col),
    )


def _parse_quoted_sheet_ref(s: _Scanner, original: str) -> AstNode:
    """Parse a quoted sheet reference like 'Sheet Name'!A1 or 'Sheet Name'!A1:B2."""
    if s.consume() != "'":
        raise FormulaParseError(original, "Expected single quote")

    sheet_chars: list[str] = []
    while True:
        ch = s.consume()
        if ch is None:
            raise FormulaParseError(original, "Unterminated quoted sheet name")
        if ch == "'":
            if s.peek() == "'":
                s.consume()
                sheet_chars.append("'")
                continue
            break
        sheet_chars.append(ch)

    sheet_name = "".join(sheet_chars)

    if s.peek() != "!":
        raise FormulaParseError(original, f"Expected '!' after quoted sheet name '{sheet_name}'")
    s.consume()

    return _parse_ref_after_sheet_bang(s, original, sheet_qualifier=sheet_name)


def _parse_ident(s: _Scanner) -> str:
    return s.take_while(lambda c: c.isalnum() or c in ("_", "."))


def _bare_sheet_name(sheet_qualifier: str) -> str:
    if sheet_qualifier.startswith("'") and sheet_qualifier.endswith("'"):
        return sheet_qualifier[1:-1].replace("''", "'")
    return sheet_qualifier


def _parse_ref_after_sheet_bang(s: _Scanner, original: str, *, sheet_qualifier: str) -> AstNode:
    """Parse a reference after ``sheet!`` (cell, whole column/row, or A1 range)."""
    sheet = _bare_sheet_name(sheet_qualifier)
    s.skip_ws()
    saved = s.i
    whole_row = _try_parse_whole_row_after_bang(s, sheet)
    if whole_row is not None:
        return whole_row
    s.i = saved
    whole_col = _try_parse_whole_column(s, original, sheet)
    if whole_col is not None:
        return whole_col
    s.i = saved
    return _parse_cell_or_range(s, original, sheet)


def _try_parse_whole_row_after_bang(s: _Scanner, sheet: str) -> WholeRowNode | None:
    s.skip_ws()
    abs_row = False
    if s.peek() == "$":
        s.consume()
        abs_row = True
    row_ch = s.peek()
    if row_ch is None or not row_ch.isdigit():
        return None
    row_str = s.take_while(lambda c: c.isdigit())
    s.skip_ws()
    if s.peek() != ":":
        return None
    s.consume()
    s.skip_ws()
    if s.peek() == "$":
        s.consume()
        abs_row = True
    row2 = s.take_while(lambda c: c.isdigit())
    if row2 != row_str:
        return None
    return WholeRowNode(sheet=sheet, row=_row_axis(s, int(row_str), abs_row))


def _parse_cell_axes(s: _Scanner, original: str) -> tuple[bool, int, bool, int]:
    s.skip_ws()
    abs_col = False
    if s.peek() == "$":
        s.consume()
        abs_col = True
    col = s.take_while(lambda c: c.isalpha())
    abs_row = False
    if s.peek() == "$":
        s.consume()
        abs_row = True
    row = s.take_while(lambda c: c.isdigit())
    if not col or not row:
        raise FormulaParseError(original, f"Invalid cell coordinate at {s.i}")
    return abs_col, int(column_index_from_string(col.upper())), abs_row, int(row)


def _cell_ref_from_parsed_axes(
    s: _Scanner, sheet: str, abs_col: bool, col_index: int, abs_row: bool, row_index: int
) -> CellRef:
    return CellRef(
        sheet=sheet,
        col=_col_axis(s, col_index, abs_col),
        row=_row_axis(s, row_index, abs_row),
    )


def _parse_cell_or_range(s: _Scanner, original: str, sheet: str) -> AstNode:
    abs_col, col_index, abs_row, row_index = _parse_cell_axes(s, original)
    start = _cell_ref_from_parsed_axes(s, sheet, abs_col, col_index, abs_row, row_index)
    s.skip_ws()
    if s.peek() == ":":
        s.consume()
        end = _parse_range_end_ref(s, original, default_sheet=sheet)
        return RangeNode(start_ref=start, end_ref=end)
    return CellRefNode(start)


def _parse_range_end_ref(s: _Scanner, original: str, default_sheet: str) -> CellRef:
    s.skip_ws()
    if s.peek() == "'":
        s.consume()
        sheet_chars: list[str] = []
        while True:
            ch = s.consume()
            if ch is None:
                raise FormulaParseError(original, "Unterminated quoted sheet name in range end")
            if ch == "'":
                if s.peek() == "'":
                    s.consume()
                    sheet_chars.append("'")
                    continue
                break
            sheet_chars.append(ch)
        sheet_name = "".join(sheet_chars)
        if s.peek() != "!":
            raise FormulaParseError(
                original,
                f"Expected '!' after quoted sheet name '{sheet_name}' in range end",
            )
        s.consume()
        abs_col, col_index, abs_row, row_index = _parse_cell_axes(s, original)
        return _cell_ref_from_parsed_axes(s, sheet_name, abs_col, col_index, abs_row, row_index)

    start = s.i
    sheet = s.take_while(lambda c: c.isalnum() or c in ("_", "."))
    s.skip_ws()
    if s.peek() == "!":
        s.consume()
        abs_col, col_index, abs_row, row_index = _parse_cell_axes(s, original)
        return _cell_ref_from_parsed_axes(s, sheet, abs_col, col_index, abs_row, row_index)
    s.i = start
    abs_col, col_index, abs_row, row_index = _parse_cell_axes(s, original)
    return _cell_ref_from_parsed_axes(s, default_sheet, abs_col, col_index, abs_row, row_index)


def _parse_args(s: _Scanner, original: str) -> list[AstNode]:
    args: list[AstNode] = []
    s.skip_ws()
    if s.peek() == ")":
        s.consume()
        return args

    while True:
        s.skip_ws()
        ch = s.peek()
        if ch == ",":
            args.append(EmptyArgNode())
        elif ch == ")":
            args.append(EmptyArgNode())
            s.consume()
            return args
        else:
            args.append(_parse_expression(s, original, min_prec=0))

        s.skip_ws()
        ch = s.peek()
        if ch == ",":
            s.consume()
            s.skip_ws()
            if s.peek() == ")":
                args.append(EmptyArgNode())
                s.consume()
                return args
            continue
        if ch == ")":
            s.consume()
            return args
        raise FormulaParseError(original, f"Expected ',' or ')', got {ch!r}")


def _parse_string(s: _Scanner, original: str) -> StringNode:
    if s.consume() != '"':
        raise FormulaParseError(original, "Expected string")
    out: list[str] = []
    while True:
        ch = s.consume()
        if ch is None:
            raise FormulaParseError(original, "Unterminated string")
        if ch == '"':
            if s.peek() == '"':
                s.consume()
                out.append('"')
                continue
            return StringNode("".join(out))
        out.append(ch)


def _parse_number(s: _Scanner, original: str) -> NumberNode:
    start = s.i
    s.take_while(lambda c: c.isdigit() or c == ".")
    ch = s.peek()
    if ch is not None and ch in ("e", "E"):
        exp_mark = s.i
        s.consume()
        sign = s.peek()
        if sign in ("+", "-"):
            s.consume()
        exp_digits = s.take_while(lambda c: c.isdigit())
        if not exp_digits:
            s.i = exp_mark
    text = s.text[start : s.i]
    try:
        return NumberNode(float(text))
    except ValueError:
        raise FormulaParseError(original, f"Invalid number literal {text!r}") from None


def _parse_error(s: _Scanner, original: str) -> ErrorNode:
    text = s.take_while(lambda c: not c.isspace() and c not in (",", ")"))
    err = XlError.from_text(text)
    if err is None:
        raise FormulaParseError(original, f"Unknown error literal {text!r}")
    return ErrorNode(err)
