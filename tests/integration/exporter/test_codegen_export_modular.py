"""Modular CodeGenerator export: multi-file package imports and runs on disk (integration).

Writes a small generated package, subprocesses Python import/execution, and asserts
the split layout remains runnable for downstream packaging workflows.
"""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path

import pytest

from excel_grapher import DependencyGraph, FormulaEvaluator, Node
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
        "__init__.py",
        "api.py",
        "data.py",
        "internals.py",
        "runtime.py",
    }

    pkg_dir = tmp_path / "exported"
    for filename, content in files.items():
        assert "excel_evaluator" not in content
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

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


def test_codegen_generate_modules_api_uses_target_map(tmp_path: Path) -> None:
    graph = _make_graph(_make_node("S!A1", None, 1.0))
    files = CodeGenerator(graph).generate_modules(["S!A1"])
    api_py = files["api.py"]
    assert "TARGETS = {" in api_py
    assert (
        "    return {target: handler(ctx, target) for target, handler in TARGETS.items()}" in api_py
    )


def test_codegen_generate_modules_emits_empty_discovery_helpers(tmp_path: Path) -> None:
    graph = _make_graph(_make_node("Sheet1!A1", None, 10.0))
    files = CodeGenerator(graph).generate_modules(["Sheet1!A1"])

    api_py = files["api.py"]
    init_py = files["__init__.py"]
    assert "def list_setters() -> list[str]:" in api_py
    assert "def list_readers() -> list[str]:" in api_py
    assert "def list_computes() -> list[str]:" in api_py
    assert "list_setters" in init_py
    assert "list_readers" in init_py
    assert "list_computes" in init_py

    pkg_dir = tmp_path / "exported_discovery"
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_discovery")
        assert pkg.list_setters() == []
        assert pkg.list_readers() == []
        assert pkg.list_computes() == []
    finally:
        sys.path.remove(str(tmp_path))
        sys.modules.pop("exported_discovery", None)


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
    data_py = files["data.py"]
    api_py = files["api.py"]

    assert "CONSTANTS = {" in data_py
    inputs_section, _, constants_section = data_py.partition("CONSTANTS = {")
    assert "DEFAULT_INPUTS = {" in inputs_section
    assert "Sheet1!A2" in inputs_section
    assert "Sheet1!A1" not in inputs_section
    assert "Sheet1!A1" in constants_section
    assert "Sheet1!A3" in constants_section
    assert "merged.update(CONSTANTS)" in api_py


def test_codegen_generate_modules_constant_blanks(tmp_path: Path) -> None:
    graph = _make_graph(
        _make_node("Sheet1!A1", None, None),
        _make_node("Sheet1!A2", None, 7.0),
    )
    files = CodeGenerator(graph).generate_modules(
        ["Sheet1!A1", "Sheet1!A2"],
        constant_blanks=True,
    )
    data_py = files["data.py"]
    api_py = files["api.py"]

    assert "CONSTANTS = {" in data_py
    inputs_section, _, constants_section = data_py.partition("CONSTANTS = {")
    assert "Sheet1!A2" in inputs_section
    assert "Sheet1!A1" not in inputs_section
    assert "Sheet1!A1" in constants_section
    assert "merged.update(CONSTANTS)" in api_py


def test_codegen_generate_modules_has_no_ty_or_ruff_diagnostics(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[3]

    graph = _make_graph(_make_node("S!A1", None, 1.0))
    files = CodeGenerator(graph).generate_modules(["S!A1"])

    pkg_root = tmp_path / "exported"
    for filename, content in files.items():
        pkg_root.mkdir(parents=True, exist_ok=True)
        (pkg_root / filename).write_text(content, encoding="utf-8")

    ruff_fix = _run(
        ["uv", "run", "--no-sync", "ruff", "check", "--fix", str(pkg_root)],
        cwd=repo_root,
    )
    assert ruff_fix.returncode == 0, f"ruff --fix failed:\n{ruff_fix.stdout}\n{ruff_fix.stderr}"

    ty = _run(
        [
            "uv",
            "run",
            "--no-sync",
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

    ruff = _run(["uv", "run", "--no-sync", "ruff", "check", str(pkg_root)], cwd=repo_root)
    assert ruff.returncode == 0, f"ruff failed after --fix:\n{ruff.stdout}\n{ruff.stderr}"
    assert ruff.stderr.strip() == ""


def test_codegen_generate_modules_has_no_ty_diagnostics_for_xlookup(tmp_path: Path) -> None:
    """XLOOKUP should not introduce undefined runtime symbols in generated modules."""
    repo_root = Path(__file__).resolve().parents[3]

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

    pkg_root = tmp_path / "exported"
    for filename, content in files.items():
        pkg_root.mkdir(parents=True, exist_ok=True)
        (pkg_root / filename).write_text(content, encoding="utf-8")

    ty = _run(
        [
            "uv",
            "run",
            "--no-sync",
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


def test_codegen_generate_modules_has_no_ty_diagnostics_for_xludf_xlookup(
    tmp_path: Path,
) -> None:
    """``_xludf.XLOOKUP`` should normalize to the same runtime symbol as ``_xlfn``."""
    repo_root = Path(__file__).resolve().parents[3]

    graph = _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", None, "a"),
        _make_node("S!B2", None, "b"),
        _make_node("S!B3", None, "c"),
        _make_node("S!C1", "=_xludf.XLOOKUP(2,S!A1:S!A3,S!B1:S!B3)", None),
    )
    files = CodeGenerator(graph).generate_modules(["S!C1"])

    pkg_root = tmp_path / "exported_xludf"
    for filename, content in files.items():
        pkg_root.mkdir(parents=True, exist_ok=True)
        (pkg_root / filename).write_text(content, encoding="utf-8")

    internals = (pkg_root / "internals.py").read_text(encoding="utf-8")
    assert "xl__xludf_xlookup" not in internals
    assert "xl_xlookup" in internals

    ty = _run(
        [
            "uv",
            "run",
            "--no-sync",
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


def test_codegen_generate_modules_embeds_xl_abs_in_runtime(tmp_path: Path) -> None:
    """Modular export embeds ``xl_abs`` when the graph uses ``ABS``."""
    graph = _make_graph(
        _make_node("S!A1", None, -2.5),
        _make_node("S!B1", "=ABS(S!A1)", None),
    )
    files = CodeGenerator(graph).generate_modules(["S!B1"])

    runtime = files["runtime.py"]
    internals = files["internals.py"]
    assert "def xl_abs" in runtime
    assert "xl_abs" in internals

    pkg_dir = tmp_path / "exported_abs"
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_abs")
        results = pkg.compute_all()
        assert results["S!B1"] == 2.5
    finally:
        sys.path.remove(str(tmp_path))
        sys.modules.pop("exported_abs", None)
        for name in list(sys.modules):
            if name.startswith("exported_abs."):
                sys.modules.pop(name, None)


def test_codegen_generate_modules_embeds_xl_exp_in_runtime(tmp_path: Path) -> None:
    """Modular export embeds ``xl_exp`` when the graph uses ``EXP``."""
    graph = _make_graph(
        _make_node("S!A1", None, 1.0),
        _make_node("S!B1", "=EXP(S!A1)", None),
    )
    files = CodeGenerator(graph).generate_modules(["S!B1"])

    runtime = files["runtime.py"]
    internals = files["internals.py"]
    assert "def xl_exp" in runtime
    assert "xl_exp" in internals

    pkg_dir = tmp_path / "exported_exp"
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("exported_exp")
        results = pkg.compute_all()
        assert results["S!B1"] == pytest.approx(2.718281828459045)
    finally:
        sys.path.remove(str(tmp_path))
        sys.modules.pop("exported_exp", None)
        for name in list(sys.modules):
            if name.startswith("exported_exp."):
                sys.modules.pop(name, None)
