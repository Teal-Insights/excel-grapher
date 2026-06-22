"""Pytest fixtures for the Tiny DSA similarity-compression benchmark."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.fixtures.tiny_dsa.workbook import (
    TINY_DSA_GROUPS,
    TINY_DSA_TARGETS,
    TinyDsaGroup,
    build_tiny_dsa_workbook,
)


@pytest.fixture
def tiny_dsa_workbook_path(tmp_path: Path) -> Path:
    """Path to a freshly built Tiny DSA xlsx workbook."""
    path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(path)
    return path


@pytest.fixture
def tiny_dsa_groups() -> tuple[TinyDsaGroup, ...]:
    """Expected compressible groups (roots, members, parallel-family tags)."""
    return TINY_DSA_GROUPS


@pytest.fixture
def tiny_dsa_targets() -> tuple[str, ...]:
    """Graph target roots used when building the Tiny DSA dependency graph."""
    return TINY_DSA_TARGETS
