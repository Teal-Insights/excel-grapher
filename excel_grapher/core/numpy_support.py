"""Detect whether the optional NumPy accelerator (`fast` extra) is installed."""

from __future__ import annotations

try:
    import numpy as np

    HAS_NUMPY = True
except ImportError:  # pragma: no cover - exercised when the `fast` extra is absent
    np = None  # type: ignore[assignment]
    HAS_NUMPY = False

__all__ = ["HAS_NUMPY", "np"]
