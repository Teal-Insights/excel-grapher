"""Package-layer boundary guards.

Dependency rules (strictest first):

- `excel_grapher.core` must not import from `runtime`, `evaluator`,
  `exporter`, or `grapher`.
- `excel_grapher.runtime` must not import from `evaluator`, `exporter`,
  or `grapher`. It may use `core`.
- `excel_grapher.grapher` must not import from `evaluator`, `exporter`,
  or `runtime`. It may use `core`.
- `excel_grapher.evaluator` must not import from `exporter`. It may use
  `core`, `runtime`, and `grapher` (for blank-range / cycle primitives).
- `excel_grapher.exporter` is the top of the stack and may depend on any
  lower layer.
"""

from __future__ import annotations

import ast
import importlib
import pkgutil
from pathlib import Path


def _forbidden_imports_in_package(
    package_name: str,
    forbidden_prefixes: tuple[str, ...],
) -> list[tuple[str, str]]:
    repo_root = Path(__file__).resolve().parents[2]
    pkg_path = repo_root / Path(*package_name.split("."))
    offenders: list[tuple[str, str]] = []
    for path in sorted(pkg_path.rglob("*.py")):
        rel = path.relative_to(repo_root / "excel_grapher")
        mod_name = "excel_grapher." + rel.with_suffix("").as_posix().replace("/", ".")
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module:
                if any(
                    node.module == bad or node.module.startswith(f"{bad}.")
                    for bad in forbidden_prefixes
                ):
                    offenders.append((mod_name, f"from {node.module} import ..."))
            elif isinstance(node, ast.Import):
                for alias in node.names:
                    if any(
                        alias.name == bad or alias.name.startswith(f"{bad}.")
                        for bad in forbidden_prefixes
                    ):
                        offenders.append((mod_name, f"import {alias.name}"))
    return offenders


def _leaked_imports(
    package_name: str, forbidden_prefixes: tuple[str, ...]
) -> list[tuple[str, str]]:
    pkg = importlib.import_module(package_name)
    offenders: list[tuple[str, str]] = []
    for mod_info in pkgutil.walk_packages(pkg.__path__, prefix=f"{package_name}."):
        mod = importlib.import_module(mod_info.name)
        for name, value in vars(mod).items():
            origin = getattr(value, "__module__", "") or ""
            if any(origin.startswith(bad) for bad in forbidden_prefixes):
                offenders.append((mod_info.name, f"{name} <- {origin}"))
    return offenders


def test_core_has_no_upward_deps() -> None:
    offenders = _forbidden_imports_in_package(
        "excel_grapher.core",
        (
            "excel_grapher.runtime",
            "excel_grapher.evaluator",
            "excel_grapher.exporter",
            "excel_grapher.grapher",
        ),
    )
    assert not offenders, f"core leaked imports: {offenders}"


def test_runtime_has_no_upward_deps() -> None:
    offenders = _leaked_imports(
        "excel_grapher.runtime",
        (
            "excel_grapher.evaluator",
            "excel_grapher.exporter",
            "excel_grapher.grapher",
        ),
    )
    assert not offenders, f"runtime leaked imports: {offenders}"


def test_grapher_has_no_upward_deps() -> None:
    offenders = _leaked_imports(
        "excel_grapher.grapher",
        (
            "excel_grapher.evaluator",
            "excel_grapher.exporter",
            "excel_grapher.runtime",
        ),
    )
    assert not offenders, f"grapher leaked imports: {offenders}"


def test_evaluator_does_not_import_exporter() -> None:
    offenders = _leaked_imports("excel_grapher.evaluator", ("excel_grapher.exporter",))
    assert not offenders, f"evaluator leaked imports from exporter: {offenders}"


def test_grapher_package_init_does_not_import_exporter() -> None:
    """The package `__init__` is not traversed by walk_packages; assert it stays layer-clean."""
    repo_root = Path(__file__).resolve().parents[2]
    init_path = repo_root / "excel_grapher" / "grapher" / "__init__.py"
    tree = ast.parse(init_path.read_text(encoding="utf-8"))
    offenders: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module:
            if node.module == "excel_grapher.exporter" or node.module.startswith(
                "excel_grapher.exporter."
            ):
                offenders.append(f"from {node.module} import ...")
        elif isinstance(node, ast.Import):
            for alias in node.names:
                name = alias.name
                if name == "excel_grapher.exporter" or name.startswith("excel_grapher.exporter."):
                    offenders.append(f"import {name}")
    assert not offenders, f"grapher/__init__.py must not import exporter: {offenders}"
