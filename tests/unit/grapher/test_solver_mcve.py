"""Load and replay `may_cycle_solver_mcve` fragments (#533)."""

from __future__ import annotations

import json
from pathlib import Path

from excel_grapher.grapher.guard import CellRef, Compare, Literal, Not
from excel_grapher.grapher.solver_mcve import load_solver_mcve, solver_mcve_from_mapping


def _tiny_fragment() -> dict[str, object]:
    c78 = "'Input 5 - Local-debt Financing'!C78"
    x4 = "lookup!X4"
    c120 = "Ext_Debt_Data!C120"
    not_c78 = {
        "type": "not",
        "operand": {
            "type": "cmp",
            "op": "=",
            "left": {"type": "cell", "key": c78},
            "right": {"type": "lit", "value": {"t": "int", "v": 0}},
        },
    }
    eq_residency = {
        "type": "cmp",
        "op": "=",
        "left": {"type": "cell", "key": c120},
        "right": {"type": "cell", "key": x4},
    }
    return {
        "schema_version": "1.0.0",
        "kind": "may_cycle_solver_mcve",
        "scc_index": 0,
        "seed": "Sheet1!A1",
        "scc_members": ["Sheet1!A1", "Sheet1!B1"],
        "guard_refs": [c78, c120, x4],
        "guard_cone_outside_scc": [c78, c120, x4, "translation!C90", "'Input 1 - Basics'!C33"],
        "missing_leaf_constraints": [],
        "feasible_cycle_path": ["Sheet1!A1", "Sheet1!B1", "Sheet1!A1"],
        "leaf_constraints": {
            c78: {
                "value": 0,
                "constraint": {
                    "kind": "annotated",
                    "base": "int | None",
                    "meta": [{"type": "Between", "min": 0, "max": 1}],
                },
            },
            "translation!C90": {
                "value": "Residency-based",
                "constraint": {"kind": "literal", "values": ["Residency-based"]},
            },
            "'Input 1 - Basics'!C33": {
                "value": "Residency-based",
                "constraint": {
                    "kind": "literal",
                    "values": ["Residency-based", "Currency-based"],
                },
            },
        },
        "nodes": {
            "Sheet1!A1": {
                "in_scc": True,
                "in_guard_cone": False,
                "is_leaf": False,
                "normalized_formula": "=Sheet1!B1",
                "value": 1,
            },
            "Sheet1!B1": {
                "in_scc": True,
                "in_guard_cone": False,
                "is_leaf": False,
                "normalized_formula": "=Sheet1!A1",
                "value": 1,
            },
            c78: {
                "in_scc": False,
                "in_guard_cone": True,
                "is_leaf": True,
                "normalized_formula": None,
                "value": 0,
            },
            c120: {
                "in_scc": False,
                "in_guard_cone": True,
                "is_leaf": False,
                "normalized_formula": "='Input 1 - Basics'!C33",
                "value": "Residency-based",
            },
            x4: {
                "in_scc": False,
                "in_guard_cone": True,
                "is_leaf": False,
                "normalized_formula": "=translation!C90",
                "value": "Residency-based",
            },
            "translation!C90": {
                "in_scc": False,
                "in_guard_cone": True,
                "is_leaf": True,
                "normalized_formula": None,
                "value": "Residency-based",
            },
            "'Input 1 - Basics'!C33": {
                "in_scc": False,
                "in_guard_cone": True,
                "is_leaf": True,
                "normalized_formula": None,
                "value": "Residency-based",
            },
        },
        "intra_scc_edges": [
            {"from": "Sheet1!A1", "to": "Sheet1!B1", "guard": not_c78},
            {"from": "Sheet1!B1", "to": "Sheet1!A1", "guard": eq_residency},
        ],
        "guard_cone_edges": [
            {"from": c120, "to": "'Input 1 - Basics'!C33", "guard": None},
            {"from": x4, "to": "translation!C90", "guard": None},
        ],
    }


def test_load_solver_mcve_from_json_file(tmp_path: Path) -> None:
    path = tmp_path / "fragment.json"
    path.write_text(json.dumps(_tiny_fragment()), encoding="utf-8")
    mcve = load_solver_mcve(path)
    assert mcve.kind == "may_cycle_solver_mcve"
    assert mcve.scc_members == ("Sheet1!A1", "Sheet1!B1")
    assert len(mcve.intra_scc_edges) == 2
    assert mcve.intra_scc_edges[0].guard == Not(
        Compare(left=CellRef("'Input 5 - Local-debt Financing'!C78"), op="=", right=Literal(0))
    )


def test_published_witness_shape_is_infeasible_with_cell_cell_equality() -> None:
    """P and NOT(P) on Ext_Debt_Data!C120 vs lookup!X4 must not be jointly feasible."""
    fragment = _tiny_fragment()
    fragment["intra_scc_edges"] = [
        {
            "from": "Sheet1!A1",
            "to": "Sheet1!B1",
            "guard": {
                "type": "not",
                "operand": {
                    "type": "cmp",
                    "op": "=",
                    "left": {"type": "cell", "key": "Ext_Debt_Data!C120"},
                    "right": {"type": "cell", "key": "lookup!X4"},
                },
            },
        },
        {
            "from": "Sheet1!B1",
            "to": "Sheet1!A1",
            "guard": {
                "type": "cmp",
                "op": "=",
                "left": {"type": "cell", "key": "Ext_Debt_Data!C120"},
                "right": {"type": "cell", "key": "lookup!X4"},
            },
        },
    ]
    mcve = solver_mcve_from_mapping(fragment)
    graph = mcve.to_graph()
    report = graph.cycle_report()
    assert report.has_must_cycles is False
    assert report.has_may_cycles is False


def test_cached_values_make_not_c78_eq_0_false() -> None:
    from excel_grapher.grapher.guard import evaluate_guard

    mcve = solver_mcve_from_mapping(_tiny_fragment())
    values = {key: rec.value for key, rec in mcve.nodes.items()}
    guard = mcve.intra_scc_edges[0].guard
    assert evaluate_guard(guard, values) is False
