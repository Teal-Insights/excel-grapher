"""Raise-based error channel for exported code (#315, #326).

Exported code raises `XlErrorException` where Excel displays an error code;
the evaluator keeps `XlError` sentinels. Parity is asserted on matching codes.
Direct ``cell_*`` calls and ``xl_eval`` must agree on top-level runtime errors.
"""

from __future__ import annotations

import ast as py_ast
from typing import Any, cast

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.types import XlErrorException
from excel_grapher.evaluator.types import XlError
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.embed import emit_runtime
from excel_grapher.exporter.export_runtime.offset import xl_index_ref
from excel_grapher.exporter.export_runtime.operators import xl_bool, xl_int
from tests.integration.utils.parity_harness import assert_codegen_matches_evaluator


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


def _compute_all(graph: DependencyGraph, targets: list[str]) -> tuple[Any, dict[str, Any]]:
    code = CodeGenerator(graph).generate(targets)
    ns: dict[str, Any] = {}
    exec(code, ns)
    return ns["compute_all"], ns


def _cell_function_sources(code: str) -> dict[str, str]:
    module = py_ast.parse(code)
    out: dict[str, str] = {}
    for node in module.body:
        if isinstance(node, py_ast.FunctionDef) and node.name.startswith("cell_"):
            source = py_ast.get_source_segment(code, node)
            assert source is not None
            out[node.name] = source
    return out


def _address_to_cell_fn(address: str) -> str:
    name: list[str] = []
    prev_underscore = False
    for ch in address.lower():
        if ch == "'":
            continue
        if "a" <= ch <= "z" or "0" <= ch <= "9":
            name.append(ch)
            prev_underscore = False
        elif not prev_underscore:
            name.append("_")
            prev_underscore = True
    base = "".join(name).strip("_")
    return f"cell_{base}"


def _assert_direct_call_matches_xl_eval(
    graph: DependencyGraph,
    target: str,
    expected: XlError,
) -> None:
    code = CodeGenerator(graph).generate([target])
    ns: dict[str, Any] = {}
    exec(code, ns)
    ctx = ns["make_context"]()
    cell_fn = ns[_address_to_cell_fn(target)]
    xl_error_exception = cast("type[BaseException]", ns["XlErrorException"])
    with pytest.raises(xl_error_exception) as direct_exc:
        cell_fn(ctx)
    with pytest.raises(xl_error_exception) as eval_exc:
        ns["xl_eval"](ctx, target, cell_fn)
    assert cast(Any, direct_exc.value).code == expected
    assert cast(Any, eval_exc.value).code == expected


def _has_xlerror_isinstance(function_source: str) -> bool:
    function = py_ast.parse(function_source).body[0]
    assert isinstance(function, py_ast.FunctionDef)
    for node in py_ast.walk(function):
        if not isinstance(node, py_ast.Call):
            continue
        if not isinstance(node.func, py_ast.Name) or node.func.id != "isinstance":
            continue
        if len(node.args) < 2:
            continue
        type_arg = node.args[1]
        if isinstance(type_arg, py_ast.Name) and type_arg.id == "XlError":
            return True
        if isinstance(type_arg, py_ast.Tuple):
            for elt in type_arg.elts:
                if isinstance(elt, py_ast.Name) and elt.id == "XlError":
                    return True
    return False


class TestExportRuntimeBoundaryHelpers:
    """Export-owned coercion and reference helpers raise at the codegen boundary."""

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


class TestGeneratedBoundaryInvariant:
    """Generated formula cells do not inspect `XlError` sentinels directly."""

    def test_if_and_choose_use_raise_only_coercion_wrappers(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", '=IF("nope", 1, 2)', None),
            _make_node("S!A2", '=CHOOSE("nope", 1, 2)', None),
        )
        code = CodeGenerator(graph).generate(["S!A1", "S!A2"])
        cells = _cell_function_sources(code)

        assert "xl_bool(" in cells["cell_s_a1"]
        assert "to_bool(" not in cells["cell_s_a1"]
        assert "xl_int(" in cells["cell_s_a2"]
        assert "to_int(" not in cells["cell_s_a2"]
        assert not any(_has_xlerror_isinstance(source) for source in cells.values())


class TestDirectCallBoundaryEquivalence:
    """Direct ``cell_*`` calls match ``xl_eval`` for top-level runtime errors (#326)."""

    @pytest.mark.parametrize(
        ("formula", "expected", "extra_nodes"),
        [
            (
                '=AVERAGEIF(S!B1:S!B2, ">5", S!C1:C3)',
                XlError.VALUE,
                (
                    _make_node("S!B1", None, 1),
                    _make_node("S!B2", None, 2),
                    _make_node("S!C1", None, 10),
                    _make_node("S!C2", None, 20),
                    _make_node("S!C3", None, 30),
                ),
            ),
            (
                "=SUM(S!B1:S!B3)",
                XlError.DIV,
                (
                    _make_node("S!B1", None, 1),
                    _make_node("S!B2", None, 2),
                    _make_node("S!B3", "=1/0", None),
                ),
            ),
            (
                '=VALUE("not-a-number")',
                XlError.VALUE,
                (),
            ),
            (
                "=OFFSET(S!B1, -5, 0)",
                XlError.REF,
                (_make_node("S!B1", None, 1),),
            ),
        ],
        ids=["averageif", "sum", "value", "offset"],
    )
    def test_direct_call_matches_xl_eval(
        self,
        formula: str,
        expected: XlError,
        extra_nodes: tuple[Node, ...],
    ) -> None:
        graph = _make_graph(*extra_nodes, _make_node("S!A1", formula, None))
        _assert_direct_call_matches_xl_eval(graph, "S!A1", expected)

    def test_embedded_runtime_uses_core_delegation_wrappers(self) -> None:
        graph = _make_graph(_make_node("S!A1", '=AVERAGEIF(S!B1:S!B2, ">5", S!C1:C3)', None))
        code = CodeGenerator(graph).generate(["S!A1"])
        assert "def averageif_cells(" in code
        assert "averageif_cells(" in code
        assert "raise_if_sentinel_float(" in code
        assert "return xl_averageif(" in _cell_function_sources(code)["cell_s_a1"]


class TestComputeAllRaises:
    """compute_all surfaces Excel errors as raised exceptions (public boundary)."""

    def test_division_by_zero_raises_div(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=1/0", None))
        compute_all, ns = _compute_all(graph, ["S!A1"])
        with pytest.raises(cast("type[BaseException]", ns["XlErrorException"])) as exc_info:
            compute_all()
        assert cast(Any, exc_info.value).code == XlError.DIV

    def test_raised_error_is_project_exception_type(self) -> None:
        """The embedded exception type mirrors the package's XlErrorException."""
        graph = _make_graph(_make_node("S!A1", "=1/0", None))
        compute_all, ns = _compute_all(graph, ["S!A1"])
        try:
            compute_all()
        except Exception as exc:
            assert type(exc).__name__ == XlErrorException.__name__
            assert str(exc) == "#DIV/0!"
        else:
            raise AssertionError("Expected compute_all to raise")

    def test_error_leaf_input_raises_on_read(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", None, XlError.NAME),
            _make_node("S!B1", "=S!A1", None),
        )
        compute_all, ns = _compute_all(graph, ["S!B1"])
        with pytest.raises(cast("type[BaseException]", ns["XlErrorException"])) as exc_info:
            compute_all()
        assert cast(Any, exc_info.value).code == XlError.NAME

    def test_error_code_is_cached_after_first_raise(self) -> None:
        """Re-reading a raising cell raises again without re-evaluating."""
        graph = _make_graph(
            _make_node("S!A1", "=1/0", None),
            _make_node("S!B1", "=S!A1", None),
            _make_node("S!C1", "=S!A1", None),
        )
        code = CodeGenerator(graph).generate(["S!B1", "S!C1"])
        ns: dict[str, Any] = {}
        exec(code, ns)
        ctx = ns["make_context"]()
        for target in ("S!B1", "S!C1"):
            with pytest.raises(cast("type[BaseException]", ns["XlErrorException"])):
                ns["xl_cell"](ctx, target)
        assert ctx.cache["S!A1"] == XlError.DIV
        assert ctx.cache["S!B1"] == XlError.DIV
        assert ctx.cache["S!C1"] == XlError.DIV


class TestErrorConsumers:
    """IFERROR/IFNA/IS* consume both raised errors and sentinel returns."""

    def test_iferror_catches_raised_division_error(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IFERROR(1/0, 99)", None))
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == 99

    def test_iferror_passes_through_success_value(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IFERROR(1+1, 99)", None))
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == 2

    def test_iferror_fallback_is_lazy(self) -> None:
        """The fallback branch is not evaluated when the value succeeds."""
        graph = _make_graph(
            _make_node("S!A1", None, 5),
            _make_node("S!B1", "=IFERROR(S!A1, 1/0)", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!B1"])
        assert result.generated_results["S!B1"] == 5

    def test_ifna_catches_na_only(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=IFNA(NA(), 7)", None),
            _make_node("S!A2", '=IFNA(MATCH(9, S!B1:S!B2, 0), "missing")', None),
            _make_node("S!B1", None, 1),
            _make_node("S!B2", None, 2),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2"])
        assert result.generated_results["S!A1"] == 7
        assert result.generated_results["S!A2"] == "missing"

    def test_ifna_reraises_other_errors(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IFNA(1/0, 7)", None))
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == XlError.DIV

    def test_iserror_true_for_raised_error(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=ISERROR(1/0)", None),
            _make_node("S!A2", "=ISERROR(1+1)", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2"])
        assert result.generated_results["S!A1"] is True
        assert result.generated_results["S!A2"] is False

    def test_isna_distinguishes_error_codes(self) -> None:
        graph = _make_graph(
            _make_node("S!A1", "=ISNA(NA())", None),
            _make_node("S!A2", "=ISNA(1/0)", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2"])
        assert result.generated_results["S!A1"] is True
        assert result.generated_results["S!A2"] is False

    def test_isnumber_propagates_argument_errors_like_evaluator(self) -> None:
        """Generic IS functions follow the evaluator's argument error precheck."""
        graph = _make_graph(
            _make_node("S!A1", "=ISNUMBER(1/0)", None),
            _make_node("S!A2", "=ISNUMBER(2)", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2"])
        assert result.generated_results["S!A1"] == XlError.DIV
        assert result.generated_results["S!A2"] is True

    def test_generic_functions_propagate_error_returning_argument(self) -> None:
        """`NA()` returns an error sentinel; a generic consumer must propagate it.

        Unlike `1/0` (which raises), an error-returning function yields an
        `XlError` sentinel. A non-thunked generic function such as `ISNUMBER`
        or `ISTEXT` must not swallow it: the evaluator's argument error
        precheck returns the error, so the export path must too.
        """
        graph = _make_graph(
            _make_node("S!A1", "=ISNUMBER(NA())", None),
            _make_node("S!A2", "=ISTEXT(NA())", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1", "S!A2"])
        assert result.generated_results["S!A1"] == XlError.NA
        assert result.generated_results["S!A2"] == XlError.NA

    def test_if_with_erroring_condition_propagates(self) -> None:
        graph = _make_graph(_make_node("S!A1", "=IF(1/0, 1, 2)", None))
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == XlError.DIV


class TestErrorCodeParity:
    """Excel error code E in the evaluator == raised XlErrorException(E) in export."""

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
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == expected

    def test_match_skips_error_cells_before_match(self) -> None:
        """Lookup scans keep Excel skip semantics over error cells."""
        graph = _make_graph(
            _make_node("S!B1", None, 1),
            _make_node("S!B2", "=1/0", None),
            _make_node("S!B3", None, 5),
            _make_node("S!A1", "=MATCH(5, S!B1:S!B3, 0)", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == 3

    def test_countif_over_error_range_matches_evaluator_precheck(self) -> None:
        """Generic functions propagate range argument errors like the evaluator."""
        graph = _make_graph(
            _make_node("S!B1", None, 10),
            _make_node("S!B2", "=1/0", None),
            _make_node("S!B3", None, 20),
            _make_node("S!A1", '=COUNTIF(S!B1:S!B3, ">5")', None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == XlError.DIV

    def test_countif_text_cells_do_not_raise_on_numeric_criteria(self) -> None:
        """Criteria coercion failures skip cells rather than raising."""
        graph = _make_graph(
            _make_node("S!B1", None, 10),
            _make_node("S!B2", None, "text"),
            _make_node("S!B3", None, 20),
            _make_node("S!A1", '=COUNTIF(S!B1:S!B3, ">5")', None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == 2

    def test_index_still_ignores_unused_error_cells(self) -> None:
        """The raise flip preserves lazy selective access from #314."""
        graph = _make_graph(
            _make_node("S!B1", None, 10),
            _make_node("S!B2", "=1/0", None),
            _make_node("S!B3", None, 30),
            _make_node("S!A1", "=INDEX(S!B1:S!B3, 3)", None),
        )
        result = assert_codegen_matches_evaluator(graph, ["S!A1"])
        assert result.generated_results["S!A1"] == 30
