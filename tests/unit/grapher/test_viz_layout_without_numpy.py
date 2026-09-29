"""Viewer layouts fail with an install hint when numpy is missing."""

from __future__ import annotations

import importlib.util

import pytest

from excel_grapher.grapher.viz_layout import clustered_force_layout


@pytest.mark.skipif(importlib.util.find_spec("numpy") is not None, reason="numpy is installed")
def test_layout_names_the_fast_extra_without_numpy() -> None:
    with pytest.raises(ImportError, match=r"excel-grapher\[viz\]"):
        clustered_force_layout(2, [(0, 1)], [0, 0])
