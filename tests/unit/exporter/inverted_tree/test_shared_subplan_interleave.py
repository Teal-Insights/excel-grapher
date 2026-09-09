"""Issue 813 — shared helpers must not splice before formula params are bound.

Consumer-set grouping can put an input-only flag in the same `_shared_*` as a
later series that needs another helper's return. Splicing at `formula_ids[0]`
then emits `helper(..., mid2=mid2)` before `mid2` is assigned.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    required_param_names,
    series_entry,
    write_workbook,
)


def _interleave_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "interleave.xlsx",
        {
            "Inputs": {"A1": 1, "B1": 2},
            "Engine": {
                "A1": "=Inputs!A1",
                "B1": "=Inputs!A1+10",
                "C1": "=Engine!B1+1",
                "D1": "=Engine!C1+Engine!A1",
                "E1": "=Engine!D1+Inputs!B1",
            },
            "Outputs": {
                "A1": "=Engine!D1",
                "B1": "=Engine!E1",
                "C1": "=Engine!C1",
            },
        },
    )


def _interleave_bindings() -> dict:
    return bindings_document(
        series_entry("x", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("y", "Inputs!B1", layout="scalar", direction="input"),
        series_entry("flag", "Engine!A1", layout="scalar", direction="internal"),
        series_entry("combo", "Engine!D1", layout="scalar", direction="internal"),
        series_entry("mid", "Engine!B1", layout="scalar", direction="internal"),
        series_entry("mid2", "Engine!C1", layout="scalar", direction="internal"),
        series_entry("shock_tail", "Engine!E1", layout="scalar", direction="internal"),
        series_entry("out_base", "Outputs!A1", layout="scalar", direction="output"),
        series_entry("out_shock", "Outputs!B1", layout="scalar", direction="output"),
        series_entry("out_mid", "Outputs!C1", layout="scalar", direction="output"),
    )


def _kwarg_uses_before_assign(source: str) -> list[tuple[str, int]]:
    """Return `(name, lineno)` for keyword values that name an unbound local."""
    tree = ast.parse(source)
    issues: list[tuple[str, int]] = []
    for fn in tree.body:
        if not isinstance(fn, ast.FunctionDef):
            continue
        bound: set[str] = {arg.arg for arg in fn.args.kwonlyargs}
        bound.update(arg.arg for arg in fn.args.args)
        for node in fn.body:
            for call in ast.walk(node):
                if not isinstance(call, ast.Call):
                    continue
                for keyword in call.keywords:
                    value = keyword.value
                    if isinstance(value, ast.Name) and value.id not in bound:
                        issues.append((value.id, value.lineno))
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name):
                        bound.add(target.id)
                    elif isinstance(target, ast.Tuple):
                        for elt in target.elts:
                            if isinstance(elt, ast.Name):
                                bound.add(elt.id)
    return issues


def test_interleaved_consumer_groups_do_not_emit_unbound_helper_params(
    tmp_path: Path,
) -> None:
    modules = generate_inverted(_interleave_workbook(tmp_path), _interleave_bindings())
    kernel = modules["api.py"]
    assert _kwarg_uses_before_assign(kernel) == []


def test_interleaved_consumer_groups_evaluate_without_unbound_local(tmp_path: Path) -> None:
    workbook = _interleave_workbook(tmp_path)
    document = _interleave_bindings()
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="issue813")
    assert required_param_names(pkg.compute_out_base) == ("x",)
    assert "y" not in all_param_names(pkg.compute_out_base)
    assert required_param_names(pkg.compute_out_mid) == ("x",)
    assert set(required_param_names(pkg.compute_out_shock)) == {"x", "y"}

    assert pkg.compute_out_base(x=1.0) == pytest.approx(13.0)
    assert pkg.compute_out_mid(x=1.0) == pytest.approx(12.0)
    assert pkg.compute_out_shock(x=1.0, y=2.0) == pytest.approx(15.0)

    _catalog, _deps, graph = inverted_graph_parts(workbook, document)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1", "Outputs!B1", "Outputs!C1"])
    assert pkg.compute_out_base(x=1.0) == pytest.approx(expected["Outputs!A1"])
    assert pkg.compute_out_shock(x=1.0, y=2.0) == pytest.approx(expected["Outputs!B1"])
    assert pkg.compute_out_mid(x=1.0) == pytest.approx(expected["Outputs!C1"])
