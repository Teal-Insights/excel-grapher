"""Load `CONSTRAINTS` tables from a `constraints.py` module.

Corpus entries (`tests/fixtures/local/corpus.toml`) and `bindings validate
--constraints` share this contract: the module exposes
`CONSTRAINTS: Mapping[str, type]` whose keys are sheet-qualified addresses.
"""

from __future__ import annotations

import importlib.util
from collections.abc import Mapping
from pathlib import Path
from types import ModuleType
from typing import Any

from excel_grapher.grapher.dynamic_refs import DynamicRefConfig


class ConstraintsLoadError(ValueError):
    """Raised when a `constraints.py` module cannot be loaded or is malformed."""


def resolve_constraints_path(workbook: Path, constraints: Path) -> Path:
    """Resolve `--constraints` against CWD, then the workbook directory.

    Args:
        workbook: Workbook path used as the relative-path fallback root.
        constraints: Explicit constraints module path from the CLI.

    Returns:
        Existing `constraints.py` path.

    Raises:
        ConstraintsLoadError: When no candidate file exists.
    """
    candidates = [constraints]
    if not constraints.is_absolute():
        candidates.append(workbook.parent / constraints)
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    tried = ", ".join(str(path) for path in candidates)
    raise ConstraintsLoadError(f"Constraints module not found: {constraints} (tried: {tried})")


def load_constraints_module(path: Path) -> ModuleType:
    """Import a constraints module and require a `CONSTRAINTS` mapping.

    Args:
        path: Filesystem path to a `constraints.py` module.

    Returns:
        The imported module (always has a mapping `CONSTRAINTS` attribute).

    Raises:
        ConstraintsLoadError: When the file is missing, cannot be imported, or
            does not expose `CONSTRAINTS` as a mapping.
    """
    resolved = path.resolve()
    if not resolved.is_file():
        raise ConstraintsLoadError(f"Constraints module not found: {resolved}")
    spec = importlib.util.spec_from_file_location(
        f"excel_grapher_constraints_{resolved.stem}",
        resolved,
    )
    if spec is None or spec.loader is None:
        raise ConstraintsLoadError(f"Cannot load constraints module: {resolved}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        raise ConstraintsLoadError(
            f"Failed to import constraints module {resolved}: {exc}"
        ) from exc
    table = getattr(module, "CONSTRAINTS", None)
    if table is None:
        raise ConstraintsLoadError(
            f"Constraints module {resolved} must define CONSTRAINTS: Mapping[str, type]"
        )
    if not isinstance(table, Mapping):
        raise ConstraintsLoadError(
            f"CONSTRAINTS in {resolved} must be a mapping, got {type(table).__name__}"
        )
    return module


def constraints_table(module: ModuleType) -> Mapping[str, Any]:
    """Return the `CONSTRAINTS` mapping from a loaded module."""
    table = getattr(module, "CONSTRAINTS", None)
    if not isinstance(table, Mapping):
        raise ConstraintsLoadError("CONSTRAINTS must be a mapping")
    return table


def dynamic_refs_from_path(path: Path) -> DynamicRefConfig:
    """Build a `DynamicRefConfig` from a `constraints.py` module.

    Args:
        path: Filesystem path to a module exposing `CONSTRAINTS`.

    Returns:
        Config built with `DynamicRefConfig.from_constraints`.
    """
    module = load_constraints_module(path)
    return DynamicRefConfig.from_constraints(constraints_table(module), {})
