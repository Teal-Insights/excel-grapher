"""Enforce dependency rules: grapher must not import evaluator or exporter."""

from __future__ import annotations

import ast
from pathlib import Path


def test_grapher_does_not_import_evaluator_or_exporter() -> None:
    """Grapher must depend only on stdlib, third-party, and core; not evaluator or exporter."""
    repo_root = Path(__file__).resolve().parent.parent.parent
    grapher_dir = repo_root / "excel_grapher" / "grapher"
    assert grapher_dir.is_dir(), f"Expected grapher package at {grapher_dir}"

    forbidden_prefixes = ("excel_grapher.evaluator", "excel_grapher.exporter")
    for path in grapher_dir.rglob("*.py"):
        if path.name.startswith("_"):
            continue
        try:
            src = path.read_text(encoding="utf-8")
            tree = ast.parse(src, filename=str(path))
        except (SyntaxError, OSError):
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    mod = alias.name.split(".", 1)[0]
                    if mod == "excel_grapher":
                        full = alias.name
                        for prefix in forbidden_prefixes:
                            if full == prefix or full.startswith(prefix + "."):
                                raise AssertionError(
                                    f"Grapher must not import evaluator or exporter: "
                                    f"{path.relative_to(repo_root)} has 'import {alias.name}'"
                                )
            elif isinstance(node, ast.ImportFrom) and node.module:
                for prefix in forbidden_prefixes:
                    if node.module == prefix or node.module.startswith(prefix + "."):
                        raise AssertionError(
                            f"Grapher must not import from evaluator or exporter: "
                            f"{path.relative_to(repo_root)} has 'from {node.module} import ...'"
                        )
