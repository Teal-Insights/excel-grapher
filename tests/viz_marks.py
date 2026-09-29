"""Marks for tests that run graph viewer layouts (numpy is optional)."""

from __future__ import annotations

import importlib.util

import pytest

requires_numpy = pytest.mark.skipif(
    importlib.util.find_spec("numpy") is None,
    reason="graph viewer layouts need numpy (the `viz` extra)",
)
