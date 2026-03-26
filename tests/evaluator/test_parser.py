from excel_grapher.evaluator.parser import (
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    EmptyArgNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    parse,
)
from excel_grapher.evaluator.types import XlError


def test_parse_literals() -> None:
    assert parse("=1") == NumberNode(1.0)
    assert parse('="abc"') == StringNode("abc")
    assert parse("=TRUE") == BoolNode(True)
    assert parse("=false") == BoolNode(False)


def test_parse_cell_reference() -> None:
    assert parse("=Sheet1!A1") == CellRefNode("Sheet1!A1")


def test_parse_range_reference_same_sheet() -> None:
    assert parse("=S!A1:B2") == RangeNode(start="S!A1", end="S!B2")


def test_parse_function_call() -> None:
    ast = parse('=SUM(S!A1, 2, "x")')
    assert ast == FunctionCallNode(
        name="SUM",
        args=[CellRefNode("S!A1"), NumberNode(2.0), StringNode("x")],
    )


def test_parse_errors() -> None:
    result = parse("=#N/A")
    assert isinstance(result, ErrorNode)
    assert result.error == XlError.NA


# --- Operator parsing tests ---


def test_parse_binary_arithmetic() -> None:
    """Test basic arithmetic operators: +, -, *, /, ^"""
    assert parse("=1+2") == BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0))
    assert parse("=3-1") == BinaryOpNode("-", NumberNode(3.0), NumberNode(1.0))
    assert parse("=2*3") == BinaryOpNode("*", NumberNode(2.0), NumberNode(3.0))
    assert parse("=6/2") == BinaryOpNode("/", NumberNode(6.0), NumberNode(2.0))
    assert parse("=2^3") == BinaryOpNode("^", NumberNode(2.0), NumberNode(3.0))


def test_parse_comparison_operators() -> None:
    """Test comparison operators: <, >, =, <=, >=, <>"""
    assert parse("=1<2") == BinaryOpNode("<", NumberNode(1.0), NumberNode(2.0))
    assert parse("=1>2") == BinaryOpNode(">", NumberNode(1.0), NumberNode(2.0))
    assert parse("=1=2") == BinaryOpNode("=", NumberNode(1.0), NumberNode(2.0))
    assert parse("=1<=2") == BinaryOpNode("<=", NumberNode(1.0), NumberNode(2.0))
    assert parse("=1>=2") == BinaryOpNode(">=", NumberNode(1.0), NumberNode(2.0))
    assert parse("=1<>2") == BinaryOpNode("<>", NumberNode(1.0), NumberNode(2.0))


def test_parse_string_concat() -> None:
    """Test string concatenation operator: &"""
    assert parse('="a"&"b"') == BinaryOpNode("&", StringNode("a"), StringNode("b"))


def test_parse_unary_negation() -> None:
    """Test unary minus operator"""
    assert parse("=-1") == UnaryOpNode("-", NumberNode(1.0))
    assert parse("=--1") == UnaryOpNode("-", UnaryOpNode("-", NumberNode(1.0)))


def test_parse_operator_precedence() -> None:
    """Test operator precedence: ^ > * = / > + = -"""
    # 1+2*3 should parse as 1+(2*3)
    assert parse("=1+2*3") == BinaryOpNode(
        "+", NumberNode(1.0), BinaryOpNode("*", NumberNode(2.0), NumberNode(3.0))
    )
    # 2^3*4 should parse as (2^3)*4
    assert parse("=2^3*4") == BinaryOpNode(
        "*", BinaryOpNode("^", NumberNode(2.0), NumberNode(3.0)), NumberNode(4.0)
    )
    # 1+2=3 should parse as (1+2)=3 (comparison is lowest precedence)
    assert parse("=1+2=3") == BinaryOpNode(
        "=", BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0)), NumberNode(3.0)
    )


def test_parse_parentheses() -> None:
    """Test parentheses for grouping"""
    # (1+2)*3 should override default precedence
    assert parse("=(1+2)*3") == BinaryOpNode(
        "*", BinaryOpNode("+", NumberNode(1.0), NumberNode(2.0)), NumberNode(3.0)
    )


def test_parse_operators_with_cell_refs() -> None:
    """Test operators with cell references"""
    assert parse("=Sheet1!A1+Sheet1!B1") == BinaryOpNode(
        "+", CellRefNode("Sheet1!A1"), CellRefNode("Sheet1!B1")
    )


def test_parse_operators_with_functions() -> None:
    """Test operators with function calls"""
    ast = parse("=SUM(S!A1:B2)+1")
    assert ast == BinaryOpNode(
        "+",
        FunctionCallNode("SUM", [RangeNode("S!A1", "S!B2")]),
        NumberNode(1.0),
    )


def test_parse_quoted_sheet_names() -> None:
    """Test parsing sheet names with spaces (quoted with single quotes)"""
    # Single cell reference with quoted sheet name
    assert parse("='My Sheet'!A1") == CellRefNode("'My Sheet'!A1")
    # Range reference with quoted sheet name
    assert parse("='Data Sheet'!A1:B2") == RangeNode(
        "'Data Sheet'!A1", "'Data Sheet'!B2"
    )
    # Function with quoted sheet name
    ast = parse("=SUM('Input Data'!A1:A10)")
    assert ast == FunctionCallNode(
        "SUM", [RangeNode("'Input Data'!A1", "'Input Data'!A10")]
    )


def test_parse_range_with_quoted_sheet_on_both_ends() -> None:
    """Test parsing ranges where both start and end have quoted sheet names"""
    # Range with quoted sheet name on both ends (same sheet)
    assert parse("='Input 8 - SDR'!C26:'Input 8 - SDR'!W26") == RangeNode(
        "'Input 8 - SDR'!C26", "'Input 8 - SDR'!W26"
    )
    # NPV function with this type of range
    ast = parse("=NPV('Input 1'!C25,'Input 8'!C26:'Input 8'!W26)")
    assert ast == FunctionCallNode(
        "NPV",
        [
            CellRefNode("'Input 1'!C25"),
            RangeNode("'Input 8'!C26", "'Input 8'!W26"),
        ],
    )


def test_parse_omitted_arguments() -> None:
    """Test parsing function calls with omitted arguments (e.g., INDEX(A1:B2,,1))"""
    # Omitted middle argument
    assert parse("=INDEX(S!A1:B2,,1)") == FunctionCallNode(
        "INDEX", [RangeNode("S!A1", "S!B2"), EmptyArgNode(), NumberNode(1.0)]
    )
    # Omitted last argument
    assert parse("=INDEX(S!A1:B2,1,)") == FunctionCallNode(
        "INDEX", [RangeNode("S!A1", "S!B2"), NumberNode(1.0), EmptyArgNode()]
    )
    # Multiple omitted arguments
    assert parse("=INDEX(S!A1:B2,,)") == FunctionCallNode(
        "INDEX", [RangeNode("S!A1", "S!B2"), EmptyArgNode(), EmptyArgNode()]
    )
    # Omitted first argument
    assert parse("=COUNT(,1)") == FunctionCallNode(
        "COUNT", [EmptyArgNode(), NumberNode(1.0)]
    )
    # Only omitted argument
    assert parse("=COUNT()") == FunctionCallNode("COUNT", [])
    assert parse("=COUNT(,)") == FunctionCallNode("COUNT", [EmptyArgNode(), EmptyArgNode()])
