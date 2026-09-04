"""Centralized paths to test fixture directories."""

from __future__ import annotations

from pathlib import Path

TESTS_ROOT = Path(__file__).resolve().parent
FIXTURES_ROOT = TESTS_ROOT / "fixtures"
SERIES_BINDINGS_FIXTURES = FIXTURES_ROOT / "series_bindings"
OPERATORS_BASELINE_FIXTURES = FIXTURES_ROOT / "operators_baseline"
DEP_TRACKING_BASELINE_FIXTURES = FIXTURES_ROOT / "dep_tracking_baseline"
TEST_SHEETS_FIXTURES = FIXTURES_ROOT / "test_sheets"
INVERTED_TREE_TINY_DSA = FIXTURES_ROOT / "inverted_tree" / "tiny_dsa"
LOCAL_CORPUS = FIXTURES_ROOT / "local"
