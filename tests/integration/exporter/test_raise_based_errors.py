"""Raise-based export-runtime helpers and evaluator error-code coverage (#315, #326).

The evaluator keeps `XlError` sentinels. Export-runtime wrappers raise
`XlErrorException` with the same codes. Address-keyed `generate()` was removed
(#764); package export uses inverted-tree with series bindings.
"""

from __future__ import annotations

from typing import Any, cast

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.types import XlErrorException
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.embed import emit_runtime
from excel_grapher.exporter.export_runtime.offset import xl_index_ref
from excel_grapher.exporter.export_runtime.operators import xl_bool, xl_int
from tests.integration.utils.parity_harness import evaluate_targets


def _make_node(address: str, formula: str | None, value: object) -> Node:
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
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


class TestExportRuntimeBoundaryHelpers:
    """Export-owned coercion and reference helpers raise at the wrapper boundary."""

    @pytest.mark.parametrize(
        ("helper", "bad_value"),
        [
            (xl_bool, "nope"),
            (xl_int, "nope"),
        ],
    )
    def test_scalar_coercion_wrappers_raise_errors(self, helper: Any, bad_value: object) -> None:
        with pytest.raises(XlErrorException) as exc_info:
            helper(bad_value)
        assert exc_info.value.code == XlError.VALUE

    def test_index_ref_raises_reference_errors(self) -> None:
        with pytest.raises(XlErrorException) as exc_info:
            xl_index_ref(("Sheet1", 1, 1, 3, 1), 99, None)
        assert exc_info.value.code == XlError.REF

    def test_offset_ref_raises_reference_errors(self) -> None:
        code = emit_runtime({"xl_offset_ref"}, include_offset_table=False)
        ns: dict[str, Any] = {}
        exec(code, ns)
        with pytest.raises(cast("type[BaseException]", ns["XlErrorException"])) as exc_info:
            ns["xl_offset_ref"](("Sheet1", 1, 1, 3, 1), -5, 0)
        assert cast(Any, exc_info.value).code == XlError.REF

    def test_averageif_raises_value_errors(self) -> None:
        code = emit_runtime({"xl_averageif"}, include_offset_table=False)
        ns: dict[str, Any] = {}
        exec(code, ns)
        with pytest.raises(cast("type[BaseException]", ns["XlErrorException"])) as exc_info:
            ns["xl_averageif"]([1, 2], ">5", [10, 20, 30])
        assert cast(Any, exc_info.value).code == XlError.VALUE

    def test_value_preserves_iso_date_fallback(self) -> None:
        code = emit_runtime({"xl_value"}, include_offset_table=False)
        ns: dict[str, Any] = {}
        exec(code, ns)
        assert ns["xl_value"]("2018-03-15") == 43174.0


class TestErrorConsumers:
    """IFERROR/IFNA/IS* consume error arguments with Excel IS semantics."""

    def test_iferror_catches_raised_division_error(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IFERROR(1/0, 99)", None))
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == 99

    def test_iferror_passes_through_success_value(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IFERROR(1+1, 99)", None))
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == 2

    def test_iferror_fallback_is_lazy(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, 5),
            _make_node("S!B1", "=IFERROR(S!A1, 1/0)", None),
        )
        results = evaluate_targets(graph, ["S!B1"])
        assert results["S!B1"] == 5

    def test_ifna_catches_na_only(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=IFNA(NA(), 7)", None),
            _make_node("S!A2", '=IFNA(MATCH(9, S!B1:S!B2, 0), "missing")', None),
            _make_node("S!B1", None, 1),
            _make_node("S!B2", None, 2),
        )
        results = evaluate_targets(graph, ["S!A1", "S!A2"])
        assert results["S!A1"] == 7
        assert results["S!A2"] == "missing"

    def test_ifna_reraises_other_errors(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IFNA(1/0, 7)", None))
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == XlError.DIV

    def test_iserror_true_for_raised_error(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=ISERROR(1/0)", None),
            _make_node("S!A2", "=ISERROR(1+1)", None),
        )
        results = evaluate_targets(graph, ["S!A1", "S!A2"])
        assert results["S!A1"] is True
        assert results["S!A2"] is False

    def test_isna_distinguishes_error_codes(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=ISNA(NA())", None),
            _make_node("S!A2", "=ISNA(1/0)", None),
        )
        results = evaluate_targets(graph, ["S!A1", "S!A2"])
        assert results["S!A1"] is True
        assert results["S!A2"] is False

    def test_isnumber_error_args_return_false(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=ISNUMBER(1/0)", None),
            _make_node("S!A2", "=ISNUMBER(2)", None),
        )
        results = evaluate_targets(graph, ["S!A1", "S!A2"])
        assert results["S!A1"] is False
        assert results["S!A2"] is True

    def test_is_family_error_args_return_false(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=ISNUMBER(NA())", None),
            _make_node("S!A2", "=ISTEXT(NA())", None),
            _make_node("S!A3", "=ISTEXT(1/0)", None),
            _make_node("S!A4", "=IF(ISNUMBER(1/0), 1, 0)", None),
        )
        results = evaluate_targets(graph, ["S!A1", "S!A2", "S!A3", "S!A4"])
        assert results["S!A1"] is False
        assert results["S!A2"] is False
        assert results["S!A3"] is False
        assert results["S!A4"] == 0

    def test_if_with_erroring_condition_propagates(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IF(1/0, 1, 2)", None))
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == XlError.DIV


class TestErrorCodeParity:
    """Evaluator returns Excel error sentinels for these formulas."""

    @pytest.mark.parametrize(
        ("formula", "expected"),
        [
            ("=1/0", XlError.DIV),
            ("=#N/A", XlError.NA),
            ("=NA()", XlError.NA),
            ('=NA()+"x"', XlError.NA),
            ('="abc"+1', XlError.VALUE),
            ("=INDEX(S!B1:S!B2, 5)", XlError.REF),
            ("=MATCH(9, S!B1:S!B2, 0)", XlError.NA),
            ("=OFFSET(S!B1, -5, 0)", XlError.REF),
            ("=SUM(S!B1:S!B3)", XlError.DIV),
            ("=CHOOSE(9, 1, 2)", XlError.VALUE),
            ('=IF("nope", 1, 2)', XlError.VALUE),
        ],
    )
    def test_error_code_parity(self, formula: str, expected: XlError) -> None:
        graph = _make_graph(
            _make_node("S!B1", None, 1),
            _make_node("S!B2", None, 2),
            _make_node("S!B3", "=1/0", None),
            _make_node("S!A1", formula, None),
        )
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == expected

    def test_match_skips_error_cells_before_match(self) -> None:
        graph = _make_graph(
            _make_node("S!B1", None, 1),
            _make_node("S!B2", "=1/0", None),
            _make_node("S!B3", None, 5),
            _make_node("S!A1", "=MATCH(5, S!B1:S!B3, 0)", None),
        )
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == 3

    def test_countif_over_error_range_skips_error_cells(self) -> None:
        graph = _make_graph(
            _make_node("S!B1", None, 10),
            _make_node("S!B2", "=1/0", None),
            _make_node("S!B3", None, 20),
            _make_node("S!A1", '=COUNTIF(S!B1:S!B3, ">5")', None),
        )
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == 2

    def test_countif_text_cells_do_not_raise_on_numeric_criteria(self) -> None:
        graph = _make_graph(
            _make_node("S!B1", None, 10),
            _make_node("S!B2", None, "text"),
            _make_node("S!B3", None, 20),
            _make_node("S!A1", '=COUNTIF(S!B1:S!B3, ">5")', None),
        )
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == 2

    def test_index_still_ignores_unused_error_cells(self) -> None:
        """INDEX does not evaluate unused cells in the lookup range."""
        graph = _make_graph(
            _make_node("S!B1", None, 10),
            _make_node("S!B2", "=1/0", None),
            _make_node("S!B3", None, 30),
            _make_node("S!A1", "=INDEX(S!B1:S!B3, 3)", None),
        )
        results = evaluate_targets(graph, ["S!A1"])
        assert results["S!A1"] == 30
