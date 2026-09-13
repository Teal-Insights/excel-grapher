"""Comparison policy of the named-package differential against the evaluator."""

from __future__ import annotations

import importlib.util
from pathlib import Path

_SCRIPT = Path(__file__).resolve().parents[3] / "scripts" / "named_differential.py"
_spec = importlib.util.spec_from_file_location("named_differential", _SCRIPT)
assert _spec is not None and _spec.loader is not None
differential = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(differential)


def test_numbers_match_within_tolerance_and_errors_must_agree() -> None:
    assert differential.classify(1.0, 1.0 + 5e-7) == "match"
    assert differential.classify(1.0, 1.001) == "mismatch"
    assert differential.classify(0.0, None) == "match"
    assert differential.classify("#DIV/0!", "#DIV/0!") == "matched_error"
    assert differential.classify("#N/A", "#DIV/0!") == "mismatch"
    assert differential.classify("abc", "abc") == "match"
    assert differential.classify(True, 1) == "match"


def test_matched_errors_are_expected_only_inside_the_chart_contract() -> None:
    contract = {"formulas": {"D51": "=IF(x, NA(), y)"}}
    assert differential.expected_error("'Chart Data'!D51", "#N/A", contract)
    assert not differential.expected_error("'Chart Data'!D52", "#N/A", contract)
    assert not differential.expected_error("'Chart Data'!D51", "#DIV/0!", contract)
