"""Package import direction for the semantic catalog contract (#722)."""

from __future__ import annotations

import ast
from pathlib import Path

from excel_grapher.semantic_model.ownership import (
    EMITTER_OWNED_MODULES,
    SEMANTIC_CONSUMER_MODULES,
)

_REPO_ROOT = Path(__file__).resolve().parents[3]
_PACKAGE_ROOT = _REPO_ROOT / "excel_grapher"


def _iter_python_files(package: str) -> list[Path]:
    directory = _PACKAGE_ROOT.joinpath(*package.split("."))
    return [path for path in directory.rglob("*.py") if not path.name.startswith(".")]


def _imported_modules(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    found: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            found.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            found.add(node.module)
    return found


def _assert_no_imports(package: str, forbidden_prefixes: tuple[str, ...]) -> None:
    for path in _iter_python_files(package):
        for module in _imported_modules(path):
            for prefix in forbidden_prefixes:
                if module == prefix or module.startswith(prefix + "."):
                    raise AssertionError(
                        f"{path.relative_to(_REPO_ROOT)} imports {module!r} "
                        f"(forbidden prefix {prefix!r})"
                    )


def test_semantic_model_does_not_import_grapher_exporter_or_evaluator() -> None:
    _assert_no_imports(
        "semantic_model",
        (
            "excel_grapher.grapher",
            "excel_grapher.exporter",
            "excel_grapher.evaluator",
        ),
    )


def test_series_bindings_does_not_import_semantic_model() -> None:
    _assert_no_imports("series_bindings", ("excel_grapher.semantic_model",))


def test_semantic_consumers_do_not_import_emitter_owned_modules() -> None:
    forbidden = tuple(sorted(EMITTER_OWNED_MODULES))
    for module in SEMANTIC_CONSUMER_MODULES:
        path = _REPO_ROOT / Path(*module.split(".")).with_suffix(".py")
        imported = _imported_modules(path)
        leaked = sorted(
            item
            for item in imported
            if item in EMITTER_OWNED_MODULES
            or any(item.startswith(prefix + ".") for prefix in forbidden)
        )
        assert not leaked, f"{module} imports emitter-owned modules: {leaked}"
