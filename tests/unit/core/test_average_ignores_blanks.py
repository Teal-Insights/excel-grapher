"""AVERAGE must ignore blanks / empty text like Excel (LIC-DSF Chart Data drift).

LIC-DSF ``Input 6 - Tailored Tests!G36`` is ``AVERAGE(G38:G41)`` where some cells
are ``=""``. Excel ignores those blanks; coercing ``None`` / ``""`` to ``0.0``
halves the average and cascades into Chart Data ratios (verify ``numeric_drift``).
"""

from __future__ import annotations

from excel_grapher.core.math_funcs import average_cells
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node


def test_average_cells_ignores_none_blank() -> None:
    assert average_cells(None, 10.0) == 10.0


def test_average_cells_ignores_empty_string() -> None:
    assert average_cells("", 10.0) == 10.0


def test_average_cells_lic_dsf_g36_shape() -> None:
    """Two numeric shocks + two empty strings → average of the numbers only."""
    shock = -0.2168778144226342
    assert average_cells("", shock, "", shock) == shock


def _lic_dsf_average_blank_graph() -> DependencyGraph:
    """Minimal Input-6 style graph: ``G36 = AVERAGE(G38:G41)`` with ``=""`` holes."""
    g = DependencyGraph()
    g.sheet_order = ["Sheet1"]
    shock = -0.2168778144226342
    g.add_node(
        make_cell_node(
            "Sheet1", "G", 38, formula='=""', normalized_formula='=""', is_leaf=False
        )
    )
    g.add_node(make_cell_node("Sheet1", "G", 39, value=shock, is_leaf=True))
    g.add_node(
        make_cell_node(
            "Sheet1", "G", 40, formula='=""', normalized_formula='=""', is_leaf=False
        )
    )
    g.add_node(make_cell_node("Sheet1", "G", 41, value=shock, is_leaf=True))
    g.add_node(
        make_cell_node(
            "Sheet1",
            "G",
            36,
            formula="=AVERAGE(G38:G41)",
            normalized_formula="=AVERAGE(Sheet1!G38:G41)",
            is_leaf=False,
        )
    )
    for member in ("Sheet1!G38", "Sheet1!G39", "Sheet1!G40", "Sheet1!G41"):
        g.add_edge("Sheet1!G36", member)
    return g


def test_evaluator_average_ignores_empty_string_formulas() -> None:
    shock = -0.2168778144226342
    with FormulaEvaluator(_lic_dsf_average_blank_graph()) as ev:
        assert ev.evaluate("Sheet1!G38") == ""
        assert ev.evaluate("Sheet1!G36") == shock


def test_codegen_average_ignores_empty_string_formulas() -> None:
    shock = -0.2168778144226342
    graph = _lic_dsf_average_blank_graph()
    with CodeGenerator(graph) as gen:
        code = gen.generate(targets=["Sheet1!G36"])
    ns: dict[str, object] = {}
    exec(code, ns)
    compute_all = ns["compute_all"]
    assert callable(compute_all)
    exported = compute_all()
    assert isinstance(exported, dict)
    assert exported["Sheet1!G36"] == shock
