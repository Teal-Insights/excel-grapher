"""Codegen export hygiene gaps (issues #252, #253)."""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator


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
    return subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, check=False)


@pytest.mark.xfail(
    reason="Issue #252: generated export is not Ruff-clean without --fix (I001).",
    strict=False,
)
def test_codegen_export_is_ruff_clean_without_fix(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[3]
    graph = _make_graph(_make_node("S!A1", None, 1.0))
    export_file = tmp_path / "exported_codegen.py"
    export_file.write_text(CodeGenerator(graph).generate(["S!A1"]), encoding="utf-8")
    try:
        ruff = _run(["uv", "run", "ruff", "check", str(export_file)], cwd=repo_root)
        assert ruff.returncode == 0, f"ruff failed:\n{ruff.stdout}\n{ruff.stderr}"
    finally:
        export_file.unlink(missing_ok=True)


@pytest.mark.xfail(
    reason="Issue #253: generated export is not ruff format --check clean without rewriting.",
    strict=False,
)
def test_codegen_export_is_ruff_format_clean_without_fix(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[3]
    graph = _make_graph(_make_node("S!A1", None, 1.0))
    export_file = tmp_path / "exported_codegen.py"
    export_file.write_text(CodeGenerator(graph).generate(["S!A1"]), encoding="utf-8")
    try:
        ruff_format = _run(
            ["uv", "run", "ruff", "format", "--check", str(export_file)],
            cwd=repo_root,
        )
        assert ruff_format.returncode == 0, (
            f"ruff format --check failed:\n{ruff_format.stdout}\n{ruff_format.stderr}"
        )
    finally:
        export_file.unlink(missing_ok=True)
