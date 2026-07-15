"""Detect whether the optional NumPy accelerator (`fast` extra) is installed."""

from __future__ import annotations

from importlib import import_module
from types import ModuleType

__all__ = ["HAS_NUMPY", "np"]


def _try_import_numpy() -> ModuleType | None:
    try:
        return import_module("numpy")
    except ImportError:  # pragma: no cover - exercised when the `fast` extra is absent
        return None


np: ModuleType | None = _try_import_numpy()
HAS_NUMPY: bool = np is not None
