from __future__ import annotations

from dataclasses import dataclass
from typing import TypeAlias

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
class CellRefNode:
    address: str  # Normalized: "Sheet!A1"


@dataclass(frozen=True, slots=True)
class RangeNode:
    start: str  # "Sheet!A1"
    end: str  # "Sheet!B2"


@dataclass(frozen=True, slots=True)
class FunctionCallNode:
    name: str
    args: list["AstNode"]


@dataclass(frozen=True, slots=True)
class BinaryOpNode:
    op: str
    left: "AstNode"
    right: "AstNode"


@dataclass(frozen=True, slots=True)
class UnaryOpNode:
    op: str
    operand: "AstNode"


AstNode: TypeAlias = (
    NumberNode
    | StringNode
    | BoolNode
    | ErrorNode
    | CellRefNode
    | RangeNode
    | FunctionCallNode
    | BinaryOpNode
    | UnaryOpNode
)


class _Scanner:
    def __init__(self, text: str) -> None:
        self.text = text
        self.i = 0

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


def parse(formula: str) -> AstNode:
    raw = formula.strip()
    if raw.startswith("="):
        raw = raw[1:].strip()

    s = _Scanner(raw)
    node = _parse_expression(s, formula, min_prec=0)
    s.skip_ws()
    if not s.eof():
        raise FormulaParseError(formula, f"Unexpected trailing input at {s.i}")
    return node


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

    if ch.isdigit() or ch == ".":
        return _parse_number(s, original)

    # Quoted sheet name: 'Sheet Name'!A1
    if ch == "'":
        return _parse_quoted_sheet_ref(s, original)

    if ch.isalpha() or ch in ("_",):
        ident = _parse_ident(s)
        upper = ident.upper()

        s.skip_ws()
        if s.peek() == "(":
            s.consume()  # '('
            args = _parse_args(s, original)
            return FunctionCallNode(name=upper, args=args)

        # Booleans
        if upper == "TRUE":
            return BoolNode(True)
        if upper == "FALSE":
            return BoolNode(False)

        # Cell ref or range: we already consumed the sheet name (before '!') or the whole token.
        # Rewind behavior is annoying; instead parse address tokens directly from raw start.
        # If there's an exclamation next, ident is the sheet name.
        if s.peek() == "!":
            s.consume()
            addr = _parse_cell_coord(s, original)
            start = f"{ident}!{addr}"
            s.skip_ws()
            if s.peek() == ":":
                s.consume()
                end = _parse_range_end(s, original, default_sheet=ident)
                return RangeNode(start=start, end=end)
            return CellRefNode(start)

        # Bare A1 is not supported because inputs are normalized and sheet-qualified.
        raise FormulaParseError(
            original, "Cell references must be sheet-qualified (e.g., Sheet1!A1)"
        )

    raise FormulaParseError(original, f"Unexpected character {ch!r} at {s.i}")


def _parse_quoted_sheet_ref(s: _Scanner, original: str) -> AstNode:
    """Parse a quoted sheet reference like 'Sheet Name'!A1 or 'Sheet Name'!A1:B2."""
    if s.consume() != "'":
        raise FormulaParseError(original, "Expected single quote")

    # Read until closing quote (Excel escapes quotes by doubling: '' -> ')
    sheet_chars: list[str] = []
    while True:
        ch = s.consume()
        if ch is None:
            raise FormulaParseError(original, "Unterminated quoted sheet name")
        if ch == "'":
            # Check for escaped quote ('')
            if s.peek() == "'":
                s.consume()
                sheet_chars.append("'")
                continue
            break
        sheet_chars.append(ch)

    sheet_name = "'" + "".join(sheet_chars) + "'"

    # Expect !
    if s.peek() != "!":
        raise FormulaParseError(
            original, f"Expected '!' after quoted sheet name '{sheet_name}'"
        )
    s.consume()

    # Parse cell coordinate
    addr = _parse_cell_coord(s, original)
    start = f"{sheet_name}!{addr}"

    s.skip_ws()
    if s.peek() == ":":
        s.consume()
        end = _parse_range_end(s, original, default_sheet=sheet_name)
        return RangeNode(start=start, end=end)

    return CellRefNode(start)


def _parse_ident(s: _Scanner) -> str:
    return s.take_while(lambda c: c.isalnum() or c in ("_", ".", " "))


def _parse_cell_coord(s: _Scanner, original: str) -> str:
    s.skip_ws()
    col = s.take_while(lambda c: c.isalpha())
    row = s.take_while(lambda c: c.isdigit())
    if not col or not row:
        raise FormulaParseError(original, f"Invalid cell coordinate at {s.i}")
    return f"{col.upper()}{row}"


def _parse_range_end(s: _Scanner, original: str, default_sheet: str) -> str:
    s.skip_ws()
    # End can be 'Quoted Sheet'!B2, Sheet!B2, or just B2.

    # Check for quoted sheet name
    if s.peek() == "'":
        s.consume()
        sheet_chars: list[str] = []
        while True:
            ch = s.consume()
            if ch is None:
                raise FormulaParseError(
                    original, "Unterminated quoted sheet name in range end"
                )
            if ch == "'":
                if s.peek() == "'":
                    s.consume()
                    sheet_chars.append("'")
                    continue
                break
            sheet_chars.append(ch)
        sheet_name = "'" + "".join(sheet_chars) + "'"
        if s.peek() != "!":
            raise FormulaParseError(
                original,
                f"Expected '!' after quoted sheet name '{sheet_name}' in range end",
            )
        s.consume()
        addr = _parse_cell_coord(s, original)
        return f"{sheet_name}!{addr}"

    # Check for unquoted sheet name
    start = s.i
    sheet = s.take_while(lambda c: c.isalnum() or c in ("_", ".", " "))
    if s.peek() == "!":
        s.consume()
        addr = _parse_cell_coord(s, original)
        return f"{sheet}!{addr}"
    # No explicit sheet: interpret the consumed token as a column part and continue digits if needed.
    s.i = start
    addr = _parse_cell_coord(s, original)
    return f"{default_sheet}!{addr}"


def _parse_args(s: _Scanner, original: str) -> list[AstNode]:
    args: list[AstNode] = []
    s.skip_ws()
    if s.peek() == ")":
        s.consume()
        return args

    while True:
        args.append(_parse_expression(s, original, min_prec=0))
        s.skip_ws()
        ch = s.peek()
        if ch == ",":
            s.consume()
            s.skip_ws()
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
            # Excel escapes quotes by doubling.
            if s.peek() == '"':
                s.consume()
                out.append('"')
                continue
            return StringNode("".join(out))
        out.append(ch)


def _parse_number(s: _Scanner, original: str) -> NumberNode:
    text = s.take_while(lambda c: c.isdigit() or c == ".")
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

