"""Shared pytest fixtures for test paths."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.paths import (
    DEP_TRACKING_BASELINE_FIXTURES,
    FIXTURES_ROOT,
    OPERATORS_BASELINE_FIXTURES,
    SERIES_BINDINGS_FIXTURES,
    TEST_SHEETS_FIXTURES,
    TESTS_ROOT,
)


@pytest.fixture(scope="session")
def tests_root() -> Path:
    return TESTS_ROOT


@pytest.fixture(scope="session")
def fixtures_root() -> Path:
    return FIXTURES_ROOT


@pytest.fixture(scope="session")
def series_bindings_fixtures() -> Path:
    return SERIES_BINDINGS_FIXTURES


@pytest.fixture(scope="session")
def operators_baseline_fixtures() -> Path:
    return OPERATORS_BASELINE_FIXTURES


@pytest.fixture(scope="session")
def dep_tracking_baseline_fixtures() -> Path:
    return DEP_TRACKING_BASELINE_FIXTURES


@pytest.fixture(scope="session")
def test_sheets_fixtures() -> Path:
    return TEST_SHEETS_FIXTURES
