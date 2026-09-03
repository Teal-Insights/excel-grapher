"""Fixtures for inverted-tree shape tests."""

from __future__ import annotations

import pytest


@pytest.fixture(params=["horizontal", "vertical"])
def orientation(request: pytest.FixtureRequest) -> str:
    """Run a shape test in both spreadsheet orientations."""
    return str(request.param)
