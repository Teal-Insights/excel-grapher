"""Verify CodeGenerator can export a small multi-file package that runs cleanly."""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
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


def test_codegen_generate_modules_executes_and_matches_evaluator(tmp_path: Path) -> None:
    graph = _make_graph(
        _make_node("Sheet1!A1", None, 10.0),
        _make_node("Sheet1!B1", "=Sheet1!A1*2", None),
    )
    targets = ["Sheet1!B1"]

    files = CodeGenerator(graph).generate_modules(targets)
    assert set(files.keys()) == {
        "exported/__init__.py",
        "exported/constants.py",
        "exported/entrypoint.py",
        "exported/inputs.py",
        "exported/internals.py",
        "exported/runtime.py",
    }

    for relpath, content in files.items():
        assert "excel_evaluator" not in content
        out_path = tmp_path / relpath
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported")
        compute_all = pkg.compute_all
        assert callable(compute_all)
        assert isinstance(pkg.DEFAULT_INPUTS, dict)

        generated_results = compute_all()
        with FormulaEvaluator(graph) as ev:
            evaluator_results = ev.evaluate(targets)
        assert generated_results == evaluator_results
    finally:
        sys.path.remove(str(tmp_path))
        sys.modules.pop("exported", None)


def test_codegen_generate_modules_entrypoint_uses_target_map(tmp_path: Path) -> None:
    graph = _make_graph(_make_node("S!A1", None, 1.0))
    files = CodeGenerator(graph).generate_modules(["S!A1"])
    entrypoint = files["exported/entrypoint.py"]
    assert "TARGETS = {" in entrypoint
    assert (
        "    return {target: handler(ctx, target) for target, handler in TARGETS.items()}"
        in entrypoint
    )


def test_codegen_generate_modules_entrypoints_exported(tmp_path: Path) -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1*2", None),
    )
    files = CodeGenerator(graph).generate_modules(
        ["S!B1"], entrypoints={"outputs-a": ["S!B1"]}
    )
    entrypoint = files["exported/entrypoint.py"]
    init_py = files["exported/__init__.py"]
    assert "def compute_outputs_a(inputs=None, *, ctx=None):" in entrypoint
    assert "compute_outputs_a" in init_py


def test_codegen_generate_modules_entrypoints_emit_named_functions(tmp_path: Path) -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!C1", "=S!A1*3", None),
    )
    files = CodeGenerator(graph).generate_modules(
        ["S!B1"],
        entrypoints={"outputs": ["S!B1", "S!C1"], "inputs-1": ["S!A1"]},
    )
    entrypoint = files["exported/entrypoint.py"]
    init_py = files["exported/__init__.py"]
    assert "TARGETS_OUTPUTS" in entrypoint
    assert "TARGETS_INPUTS_1" in entrypoint
    assert "def compute_outputs(inputs=None, *, ctx=None):" in entrypoint
    assert "def compute_inputs_1(inputs=None, *, ctx=None):" in entrypoint
    assert "compute_outputs" in init_py
    assert "compute_inputs_1" in init_py


def test_codegen_generate_modules_splits_constants(tmp_path: Path) -> None:
    graph = _make_graph(
        _make_node("Sheet1!A1", None, 10.0),
        _make_node("Sheet1!A2", None, "hi"),
        _make_node("Sheet1!A3", None, 5.0),
    )
    files = CodeGenerator(graph).generate_modules(
        ["Sheet1!A1", "Sheet1!A2", "Sheet1!A3"],
        constant_types={"number"},
    )
    inputs_py = files["exported/inputs.py"]
    constants_py = files["exported/constants.py"]
    entrypoint_py = files["exported/entrypoint.py"]

    assert "DEFAULT_INPUTS = {" in inputs_py
    assert "Sheet1!A2" in inputs_py
    assert "Sheet1!A1" not in inputs_py
    assert "CONSTANTS = {" in constants_py
    assert "Sheet1!A1" in constants_py
    assert "Sheet1!A3" in constants_py
    assert "merged.update(CONSTANTS)" in entrypoint_py


def test_codegen_generate_modules_constant_blanks(tmp_path: Path) -> None:
    graph = _make_graph(
        _make_node("Sheet1!A1", None, None),
        _make_node("Sheet1!A2", None, 7.0),
    )
    files = CodeGenerator(graph).generate_modules(
        ["Sheet1!A1", "Sheet1!A2"],
        constant_blanks=True,
    )
    inputs_py = files["exported/inputs.py"]
    constants_py = files["exported/constants.py"]
    entrypoint_py = files["exported/entrypoint.py"]

    assert "Sheet1!A2" in inputs_py
    assert "Sheet1!A1" not in inputs_py
    assert "CONSTANTS = {" in constants_py
    assert "Sheet1!A1" in constants_py
    assert "merged.update(CONSTANTS)" in entrypoint_py


def test_codegen_generate_modules_has_no_ty_or_ruff_diagnostics(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]

    graph = _make_graph(_make_node("S!A1", None, 1.0))
    files = CodeGenerator(graph).generate_modules(["S!A1"])

    for relpath, content in files.items():
        out_path = tmp_path / relpath
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
    pkg_root = tmp_path / "exported"

    ruff_fix = _run(
        ["uv", "run", "ruff", "check", "--fix", str(pkg_root)],
        cwd=repo_root,
    )
    assert ruff_fix.returncode == 0, f"ruff --fix failed:\n{ruff_fix.stdout}\n{ruff_fix.stderr}"

    ty = _run(
        [
            "uv",
            "run",
            "ty",
            "check",
            "--project",
            str(repo_root),
            "--extra-search-path",
            str(tmp_path),
            str(pkg_root),
        ],
        cwd=repo_root,
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
    assert ty.stderr.strip() == ""

    ruff = _run(["uv", "run", "ruff", "check", str(pkg_root)], cwd=repo_root)
    assert ruff.returncode == 0, f"ruff failed after --fix:\n{ruff.stdout}\n{ruff.stderr}"
    assert ruff.stderr.strip() == ""


def test_codegen_generate_modules_has_no_ty_diagnostics_with_hyphenated_dir(
    tmp_path: Path,
) -> None:
    """Package names should be normalized to importable folder names."""
    repo_root = Path(__file__).resolve().parents[1]

    graph = _make_graph(_make_node("S!A1", None, 1.0))
    files = CodeGenerator(graph).generate_modules(
        ["S!A1"], package_name="lic-dsf-template"
    )

    for relpath, content in files.items():
        out_path = tmp_path / relpath
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
    pkg_root = tmp_path / "lic_dsf_template"

    ty = _run(
        [
            "uv",
            "run",
            "ty",
            "check",
            "--project",
            str(repo_root),
            "--extra-search-path",
            str(tmp_path),
            str(pkg_root),
        ],
        cwd=repo_root,
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
    assert ty.stderr.strip() == ""


def test_codegen_generate_modules_has_no_ty_diagnostics_for_xlookup(tmp_path: Path) -> None:
    """XLOOKUP should not introduce undefined runtime symbols in generated modules."""
    repo_root = Path(__file__).resolve().parents[1]

    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", None, "a"),
        _make_node("S!B2", None, "b"),
        _make_node("S!B3", None, "c"),
        _make_node("S!C1", "=_xlfn.XLOOKUP(2,S!A1:S!A3,S!B1:S!B3)", None),
    )
    files = CodeGenerator(graph).generate_modules(["S!C1"])

    for relpath, content in files.items():
        out_path = tmp_path / relpath
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
    pkg_root = tmp_path / "exported"

    ty = _run(
        [
            "uv",
            "run",
            "ty",
            "check",
            "--project",
            str(repo_root),
            "--extra-search-path",
            str(tmp_path),
            str(pkg_root),
        ],
        cwd=repo_root,
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"
    assert ty.stderr.strip() == ""

