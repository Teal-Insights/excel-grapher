"""Regression test for ty diagnostics in generated exports for range arguments.

This captures a class of failures where generated code passes a Python list-of-lists
for a range into functions typed as `CellValue` (e.g. xl_sum), which ty flags.
"""

from __future__ import annotations

import subprocess
from pathlib import Path

from excel_grapher import DependencyGraph
from excel_grapher import Node

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


def test_ty_check_generated_sum_range_has_no_diagnostics(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]

    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!B1", "=SUM(S!A1:S!A2)", None),
    )
    code = CodeGenerator(graph).generate(["S!B1"])

    export_file = tmp_path / "exported_sum_range_codegen.py"
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

