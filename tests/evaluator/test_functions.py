"""Tests for Excel function implementations."""

from typing import cast

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.name_utils import parse_address
from excel_grapher.evaluator.types import XlError


def _make_node(address: str, formula: str | None, value: object) -> Node:
    """Helper to create a Node from a sheet-qualified address."""
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _make_graph(*nodes: Node) -> DependencyGraph:
    """Helper to create a DependencyGraph from nodes."""
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


# --- Logic function tests ---


def test_true_false_as_function_calls() -> None:
    """TRUE() and FALSE() as zero-arg function calls should return booleans."""
    graph = _make_graph(
        _make_node("S!A1", "=TRUE()", None),
        _make_node("S!A2", "=FALSE()", None),
        _make_node("S!A3", "=IF(TRUE(), 1, 2)", None),
        _make_node("S!A4", "=VLOOKUP(1, S!B1:S!C2, 2, FALSE())", None),
        _make_node("S!B1", None, 1),
        _make_node("S!B2", None, 2),
        _make_node("S!C1", None, "found"),
        _make_node("S!C2", None, "other"),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3", "S!A4"])
        assert result["S!A1"] is True
        assert result["S!A2"] is False
        assert result["S!A3"] == 1
        assert result["S!A4"] == "found"


def test_iserror() -> None:
    """Test ISERROR function."""
    graph = _make_graph(
        _make_node("S!A1", "=ISERROR(#VALUE!)", None),
        _make_node("S!A2", "=ISERROR(#N/A)", None),
        _make_node("S!A3", "=ISERROR(1)", None),
        _make_node("S!A4", '=ISERROR("text")', None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3", "S!A4"])
        assert result["S!A1"] is True
        assert result["S!A2"] is True
        assert result["S!A3"] is False
        assert result["S!A4"] is False


def test_isna() -> None:
    """Test ISNA function."""
    graph = _make_graph(
        _make_node("S!A1", "=ISNA(#N/A)", None),
        _make_node("S!A2", "=ISNA(#VALUE!)", None),
        _make_node("S!A3", "=ISNA(1)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3"])
        assert result["S!A1"] is True
        assert result["S!A2"] is False
        assert result["S!A3"] is False


def test_and_or() -> None:
    """Test AND and OR functions."""
    graph = _make_graph(
        _make_node("S!A1", "=AND(TRUE, TRUE)", None),
        _make_node("S!A2", "=AND(TRUE, FALSE)", None),
        _make_node("S!A3", "=OR(TRUE, FALSE)", None),
        _make_node("S!A4", "=OR(FALSE, FALSE)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3", "S!A4"])
        assert result["S!A1"] is True
        assert result["S!A2"] is False
        assert result["S!A3"] is True
        assert result["S!A4"] is False


# --- Aggregation function tests ---


def test_sum() -> None:
    """Test SUM function."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", "=SUM(S!A1:A3)", None),
        _make_node("S!B2", "=SUM(1, 2, 3)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!B2"])
        assert result["S!B1"] == 6.0
        assert result["S!B2"] == 6.0


def test_average() -> None:
    """Test AVERAGE function."""
    graph = _make_graph(
        _make_node("S!A1", None, 2),
        _make_node("S!A2", None, 4),
        _make_node("S!A3", None, 6),
        _make_node("S!B1", "=AVERAGE(S!A1:A3)", None),
        _make_node("S!B2", "=AVERAGE(2, 4, 6)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!B2"])
        assert result["S!B1"] == 4.0
        assert result["S!B2"] == 4.0


def test_min_max() -> None:
    """Test MIN and MAX functions."""
    graph = _make_graph(
        _make_node("S!A1", "=MIN(1, 2, 3)", None),
        _make_node("S!A2", "=MAX(1, 2, 3)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == 1.0
        assert result["S!A2"] == 3.0


def test_count_counta() -> None:
    """Test COUNT and COUNTA functions."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, "text"),
        _make_node("S!A3", None, None),
        _make_node("S!A4", None, 2),
        _make_node("S!B1", "=COUNT(S!A1:A4)", None),
        _make_node("S!B2", "=COUNTA(S!A1:A4)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!B2"])
        assert result["S!B1"] == 2  # Only counts numbers
        assert result["S!B2"] == 3  # Counts non-empty cells


def test_sumproduct() -> None:
    """Test SUMPRODUCT function."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", None, 3),
        _make_node("S!B2", None, 4),
        _make_node("S!C1", "=SUMPRODUCT(S!A1:A2, S!B1:B2)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!C1"])
        # 1*3 + 2*4 = 3 + 8 = 11
        assert result["S!C1"] == 11.0


# --- Lookup function tests ---


def test_index() -> None:
    """Test INDEX function."""
    graph = _make_graph(
        _make_node("S!A1", None, "a"),
        _make_node("S!A2", None, "b"),
        _make_node("S!A3", None, "c"),
        _make_node("S!B1", None, 1),
        _make_node("S!B2", None, 2),
        _make_node("S!B3", None, 3),
        # INDEX(array, row, col)
        _make_node("S!C1", "=INDEX(S!A1:B3, 2, 2)", None),
        # INDEX(array, row) for single column
        _make_node("S!C2", "=INDEX(S!A1:A3, 2)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!C1", "S!C2"])
        assert result["S!C1"] == 2  # Row 2, Col 2 -> B2
        assert result["S!C2"] == "b"  # Row 2 of single column


def test_match() -> None:
    """Test MATCH function."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!A3", None, 30),
        # MATCH(lookup_value, lookup_array, match_type)
        # match_type=0: exact match
        _make_node("S!B1", "=MATCH(20, S!A1:A3, 0)", None),
        _make_node("S!B2", "=MATCH(25, S!A1:A3, 0)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!B2"])
        assert result["S!B1"] == 2  # 20 is at position 2
        assert result["S!B2"] == XlError.NA  # 25 not found


def test_choose() -> None:
    """Test CHOOSE function."""
    graph = _make_graph(
        _make_node("S!A1", '=CHOOSE(2, "a", "b", "c")', None),
        _make_node("S!A2", "=CHOOSE(1, 10, 20, 30)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == "b"  # Index 2 -> second value
        assert result["S!A2"] == 10  # Index 1 -> first value


# --- Type checking function tests ---


def test_isnumber() -> None:
    """Test ISNUMBER function."""
    graph = _make_graph(
        _make_node("S!A1", "=ISNUMBER(1)", None),
        _make_node("S!A2", '=ISNUMBER("text")', None),
        _make_node("S!A3", "=ISNUMBER(TRUE)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3"])
        assert result["S!A1"] is True
        assert result["S!A2"] is False
        assert result["S!A3"] is False  # TRUE is a boolean, not a number


def test_istext() -> None:
    """Test ISTEXT function."""
    graph = _make_graph(
        _make_node("S!A1", '=ISTEXT("hello")', None),
        _make_node("S!A2", "=ISTEXT(123)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] is True
        assert result["S!A2"] is False


def test_isblank() -> None:
    """Test ISBLANK function."""
    graph = _make_graph(
        _make_node("S!A1", None, None),  # Blank cell
        _make_node("S!A2", None, ""),  # Empty string
        _make_node("S!A3", None, 0),  # Zero
        _make_node("S!B1", "=ISBLANK(S!A1)", None),
        _make_node("S!B2", "=ISBLANK(S!A2)", None),
        _make_node("S!B3", "=ISBLANK(S!A3)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!B2", "S!B3"])
        assert result["S!B1"] is True  # None is blank
        assert result["S!B2"] is False  # Empty string is not blank in Excel
        assert result["S!B3"] is False  # Zero is not blank


# --- Reference function tests ---


def test_row_column() -> None:
    """Test ROW and COLUMN functions."""
    graph = _make_graph(
        _make_node("S!A1", "=ROW(S!B5)", None),
        _make_node("S!A2", "=COLUMN(S!C3)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == 5  # Row 5
        assert result["S!A2"] == 3  # Column C = 3


def test_columns() -> None:
    """Test COLUMNS function."""
    graph = _make_graph(
        _make_node("S!A1", "=COLUMNS(S!A1:C3)", None),
        _make_node("S!A2", "=COLUMNS(S!B1:B10)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == 3  # A to C = 3 columns
        assert result["S!A2"] == 1  # Single column


def test_na_function() -> None:
    """Test NA function."""
    graph = _make_graph(_make_node("S!A1", "=NA()", None))
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1"])
        assert result["S!A1"] == XlError.NA


# --- Rounding function tests ---


def test_round() -> None:
    """Test ROUND function."""
    graph = _make_graph(
        _make_node("S!A1", "=ROUND(2.567, 2)", None),
        _make_node("S!A2", "=ROUND(2.5, 0)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == 2.57
        assert result["S!A2"] == 2.0  # Python's banker's rounding


def test_rounddown() -> None:
    """Test ROUNDDOWN function."""
    graph = _make_graph(
        _make_node("S!A1", "=ROUNDDOWN(2.567, 2)", None),
        _make_node("S!A2", "=ROUNDDOWN(2.9, 0)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == 2.56
        assert result["S!A2"] == 2.0


# --- Text function tests ---


def test_left_right_mid() -> None:
    """Test LEFT, RIGHT, MID functions."""
    graph = _make_graph(
        _make_node("S!A1", '=LEFT("Hello", 2)', None),
        _make_node("S!A2", '=RIGHT("Hello", 2)', None),
        _make_node("S!A3", '=MID("Hello", 2, 3)', None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3"])
        assert result["S!A1"] == "He"
        assert result["S!A2"] == "lo"
        assert result["S!A3"] == "ell"


def test_concatenate() -> None:
    """Test CONCATENATE function."""
    graph = _make_graph(
        _make_node("S!A1", '=CONCATENATE("Hello", " ", "World")', None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1"])
        assert result["S!A1"] == "Hello World"


def test_text() -> None:
    """Test TEXT function."""
    graph = _make_graph(
        _make_node("S!A1", '=TEXT(1234.5, "0.00")', None),
        _make_node("S!A2", '=TEXT(0.25, "0%")', None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert result["S!A1"] == "1234.50"
        assert result["S!A2"] == "25%"


# --- Additional lookup function tests ---


def test_vlookup_exact() -> None:
    """Test VLOOKUP with exact match."""
    graph = _make_graph(
        # Table: A1:B3
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!A3", None, 30),
        _make_node("S!B1", None, "ten"),
        _make_node("S!B2", None, "twenty"),
        _make_node("S!B3", None, "thirty"),
        # VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)
        _make_node("S!C1", "=VLOOKUP(20, S!A1:B3, 2, FALSE)", None),
        _make_node("S!C2", "=VLOOKUP(25, S!A1:B3, 2, FALSE)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!C1", "S!C2"])
        assert result["S!C1"] == "twenty"
        assert result["S!C2"] == XlError.NA


def test_vlookup_approximate() -> None:
    """Test VLOOKUP with approximate match."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!A3", None, 30),
        _make_node("S!B1", None, "ten"),
        _make_node("S!B2", None, "twenty"),
        _make_node("S!B3", None, "thirty"),
        # Approximate match (TRUE or omitted)
        _make_node("S!C1", "=VLOOKUP(25, S!A1:B3, 2, TRUE)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!C1"])
        # 25 falls between 20 and 30, so it returns the value for 20
        assert result["S!C1"] == "twenty"


def test_hlookup_exact() -> None:
    """Test HLOOKUP with exact match."""
    graph = _make_graph(
        # Table: A1:C2 (horizontal)
        _make_node("S!A1", None, 10),
        _make_node("S!B1", None, 20),
        _make_node("S!C1", None, 30),
        _make_node("S!A2", None, "ten"),
        _make_node("S!B2", None, "twenty"),
        _make_node("S!C2", None, "thirty"),
        # HLOOKUP(lookup_value, table_array, row_index_num, range_lookup)
        _make_node("S!D1", "=HLOOKUP(20, S!A1:C2, 2, FALSE)", None),
        _make_node("S!D2", "=HLOOKUP(25, S!A1:C2, 2, FALSE)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!D1", "S!D2"])
        assert result["S!D1"] == "twenty"
        assert result["S!D2"] == XlError.NA


def test_hlookup_approximate() -> None:
    """Test HLOOKUP with approximate match."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", None, 20),
        _make_node("S!C1", None, 30),
        _make_node("S!A2", None, "ten"),
        _make_node("S!B2", None, "twenty"),
        _make_node("S!C2", None, "thirty"),
        _make_node("S!D1", "=HLOOKUP(25, S!A1:C2, 2, TRUE)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!D1"])
        # 25 falls between 20 and 30, returns value for 20
        assert result["S!D1"] == "twenty"


# --- Financial function tests ---


def test_npv() -> None:
    """Test NPV function."""
    graph = _make_graph(
        _make_node("S!A1", None, -10000),  # Initial investment
        _make_node("S!A2", None, 3000),
        _make_node("S!A3", None, 4200),
        _make_node("S!A4", None, 6800),
        # NPV(rate, value1, [value2], ...)
        # Note: Excel's NPV starts from period 1, so initial investment is often added separately
        _make_node("S!B1", "=NPV(0.1, S!A2:A4)+S!A1", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1"])
        # NPV = 3000/1.1 + 4200/1.1^2 + 6800/1.1^3 = 2727.27 + 3471.07 + 5109.47 = 11307.29
        # Then -10000 + 11307.29 = 1307.29
        assert abs(float(cast(int | float, result["S!B1"])) - 1307.29) < 0.1


def test_npv_simple() -> None:
    """Test NPV with simple cash flows."""
    graph = _make_graph(
        _make_node("S!A1", "=NPV(0.1, 100, 100, 100)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1"])
        # 100/1.1 + 100/1.1^2 + 100/1.1^3 = 90.91 + 82.64 + 75.13 = 248.68
        assert abs(float(cast(int | float, result["S!A1"])) - 248.68) < 0.1


# --- Statistical function tests ---


def test_stdev() -> None:
    """Test STDEV function (sample standard deviation)."""
    graph = _make_graph(
        _make_node("S!A1", None, 2),
        _make_node("S!A2", None, 4),
        _make_node("S!A3", None, 4),
        _make_node("S!A4", None, 4),
        _make_node("S!A5", None, 5),
        _make_node("S!A6", None, 5),
        _make_node("S!A7", None, 7),
        _make_node("S!A8", None, 9),
        _make_node("S!B1", "=STDEV(S!A1:A8)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1"])
        # Mean = 40/8 = 5, variance = sum((x-5)^2)/(n-1) = 36/7 ≈ 5.14
        # STDEV ≈ 2.27
        assert abs(float(cast(int | float, result["S!B1"])) - 2.138) < 0.01


def test_stdev_single_value() -> None:
    """Test STDEV returns error for single value."""
    graph = _make_graph(_make_node("S!A1", "=STDEV(5)", None))
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1"])
        assert result["S!A1"] == XlError.DIV


def test_countif_numeric_and_wildcard() -> None:
    """Test COUNTIF with numeric comparison and basic wildcards."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!A4", None, 2),
        _make_node("S!A5", None, None),
        _make_node("S!B1", '=COUNTIF(S!A1:A5, ">=2")', None),
        _make_node("S!C1", None, "Apple"),
        _make_node("S!C2", None, "apricot"),
        _make_node("S!C3", None, "banana"),
        _make_node("S!D1", '=COUNTIF(S!C1:C3, "a*")', None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!D1"])
        assert result["S!B1"] == 3
        assert result["S!D1"] == 2


def test_large() -> None:
    """Test LARGE returns kth largest value."""
    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 3),
        _make_node("S!A3", None, 2),
        _make_node("S!A4", None, 5),
        _make_node("S!B1", "=LARGE(S!A1:A4, 2)", None),
        _make_node("S!B2", "=LARGE(S!A1:A4, 99)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!B2"])
        assert result["S!B1"] == 3.0
        assert result["S!B2"] == XlError.NUM


def test_rank() -> None:
    """Test RANK for descending (default) and ascending order."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 20),
        _make_node("S!A3", None, 20),
        _make_node("S!A4", None, 40),
        _make_node("S!B1", "=RANK(20, S!A1:A4)", None),
        _make_node("S!B2", "=RANK(25, S!A1:A4)", None),
        _make_node("S!B3", "=RANK(20, S!A1:A4, 1)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!B1", "S!B2", "S!B3"])
        assert result["S!B1"] == 2
        assert result["S!B2"] == 2
        assert result["S!B3"] == 2


def test_normdist() -> None:
    """Test NORMDIST cumulative and density."""
    graph = _make_graph(
        _make_node("S!A1", "=NORMDIST(0, 0, 1, TRUE)", None),
        _make_node("S!A2", "=NORMDIST(0, 0, 1, FALSE)", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2"])
        assert abs(float(cast(int | float, result["S!A1"])) - 0.5) < 1e-12
        assert abs(float(cast(int | float, result["S!A2"])) - 0.3989422804014327) < 1e-12


def test_address() -> None:
    """Test ADDRESS builds A1 references with absolute/relative flags and sheet names."""
    graph = _make_graph(
        _make_node("S!A1", "=ADDRESS(5, 3)", None),
        _make_node("S!A2", "=ADDRESS(5, 3, 4)", None),
        _make_node("S!A3", '=ADDRESS(5, 3, 1, TRUE, "My Sheet")', None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!A1", "S!A2", "S!A3"])
        assert result["S!A1"] == "$C$5"
        assert result["S!A2"] == "C5"
        assert result["S!A3"] == "'My Sheet'!$C$5"


def test_offset_returns_range() -> None:
    """Test OFFSET returns a reference that can be consumed by other functions like SUM."""
    graph = _make_graph(
        _make_node("S!A1", None, None),
        _make_node("S!B2", None, 10),
        _make_node("S!B3", None, 20),
        _make_node("S!C1", "=SUM(OFFSET(S!A1, 1, 1, 2, 1))", None),
    )
    with FormulaEvaluator(graph) as ev:
        result = ev.evaluate(["S!C1"])
        assert result["S!C1"] == 30.0
