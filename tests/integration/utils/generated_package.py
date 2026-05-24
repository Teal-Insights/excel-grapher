"""Helpers for writing and importing generated exporter packages in tests."""

from __future__ import annotations

import importlib
import sys
from collections.abc import Mapping
from pathlib import Path
from types import ModuleType


def purge_module_cache(package_name: str) -> None:
    prefix = f"{package_name}."
    for name in list(sys.modules):
        if name == package_name or name.startswith(prefix):
            del sys.modules[name]


def write_generated_package(tmp_path: Path, files: Mapping[str, str]) -> None:
    for relpath, content in files.items():
        out_path = tmp_path / relpath
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")


def import_generated_package(
    tmp_path: Path,
    files: Mapping[str, str],
    *,
    package_name: str = "exported",
) -> ModuleType:
    write_generated_package(tmp_path, files)
    purge_module_cache(package_name)
    sys.path.insert(0, str(tmp_path))
    try:
        return importlib.import_module(package_name)
    except Exception:
        sys.path.remove(str(tmp_path))
        raise
