"""Regression test for ty diagnostics in generated exports.

This test is intentionally RED today: it captures a ty failure seen in large
generated exports where `int(...)` is called on a `CellValue`-shaped union.
"""

from __future__ import annotations

import subprocess
from pathlib import Path

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.codegen import CodeGenerator
from excel_grapher.evaluator.name_utils import parse_address


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


def _run(cmd: list[str], *, cwd: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        check=False,
    )


def test_ty_check_generated_choose_has_no_diagnostics(tmp_path: Path) -> None:
    """Generated code should typecheck cleanly with ty."""
    repo_root = Path(__file__).resolve().parents[1]

    graph = _make_graph(
        _make_node("S!A1", None, 2),
        _make_node("S!B1", '=CHOOSE(S!A1, "a", "b", "c")', None),
    )
    code = CodeGenerator(graph).generate(["S!B1"])

    export_file = tmp_path / "exported_choose_codegen.py"
    export_file.write_text(code, encoding="utf-8")

    try:
        ty = _run(
            ["uv", "run", "ty", "check", "--project", str(repo_root), str(export_file)],
            cwd=repo_root,
        )
        assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
        assert ty.stderr.strip() == ""
    finally:
        export_file.unlink(missing_ok=True)

