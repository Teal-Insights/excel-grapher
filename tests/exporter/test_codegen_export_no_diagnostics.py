"""Ensure CodeGenerator exports pass typechecking and linting when written to disk."""

from __future__ import annotations

import subprocess
from pathlib import Path

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.codegen import CodeGenerator
from excel_grapher.evaluator.name_utils import parse_address


def _make_node(address: str, formula: str | None, value: object) -> Node:
    """Create a Node from a sheet-qualified address."""
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
    """Create a DependencyGraph from nodes."""
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


def test_codegen_export_has_no_ty_or_ruff_diagnostics(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]

    graph = _make_graph(_make_node("S!A1", None, 1.0))
    code = CodeGenerator(graph).generate(["S!A1"])

    export_file = tmp_path / "exported_codegen.py"
    export_file.write_text(code, encoding="utf-8")

    try:
        ruff_fix = _run(
            ["uv", "run", "ruff", "check", "--fix", str(export_file)],
            cwd=repo_root,
        )
        assert ruff_fix.returncode == 0, f"ruff --fix failed:\n{ruff_fix.stdout}\n{ruff_fix.stderr}"

        ty = _run(
            ["uv", "run", "ty", "check", "--project", str(repo_root), str(export_file)],
            cwd=repo_root,
        )
        assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
        assert ty.stderr.strip() == ""

        ruff = _run(["uv", "run", "ruff", "check", str(export_file)], cwd=repo_root)
        assert ruff.returncode == 0, f"ruff failed after --fix:\n{ruff.stdout}\n{ruff.stderr}"
        assert ruff.stderr.strip() == ""
    finally:
        export_file.unlink(missing_ok=True)

