"""Tests for Excel function name normalization and codegen naming."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.core.excel_function_names import excel_func_to_python_runtime_name
from excel_grapher.core.excel_function_names import (
    normalize_excel_function_name as _normalize_excel_function_name,
)
from excel_grapher.evaluator.functions import FUNCTIONS
from excel_grapher.evaluator.name_utils import excel_func_to_python, normalize_excel_function_name

_REPO_ROOT = Path(__file__).resolve().parents[3]
_SOURCE_ROOT = _REPO_ROOT / "excel_grapher"


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("SUM", "SUM"),
        ("sum", "SUM"),
        ("_xlfn.SUM", "SUM"),
        ("_XLFN.SUM", "SUM"),
        ("IFNA", "IFNA"),
        ("_xlfn.IFNA", "IFNA"),
        ("XLOOKUP", "XLOOKUP"),
        ("_xlfn.XLOOKUP", "XLOOKUP"),
        ("NUMBERVALUE", "NUMBERVALUE"),
        ("_xlfn.NUMBERVALUE", "NUMBERVALUE"),
        ("IFS", "IFS"),
        ("_xlfn.IFS", "IFS"),
        ("SUMPRODUCT", "SUMPRODUCT"),
        ("_xlfn.SUMPRODUCT", "SUMPRODUCT"),
        ("_xludf.IFNA", "IFNA"),
        ("_xludf.XLOOKUP", "XLOOKUP"),
        ("NORM.DIST", "NORM.DIST"),
        ("_xlfn.NORM.DIST", "NORM.DIST"),
    ],
)
def test_xlfn_prefix_normalizes_to_canonical_name(raw: str, expected: str) -> None:
    from excel_grapher.evaluator.functions import FUNCTIONS

    assert normalize_excel_function_name(raw) == expected
    assert _normalize_excel_function_name(raw, registered_builtins=frozenset(FUNCTIONS)) == expected


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("_xludf.MYADDIN", "_XLUDF.MYADDIN"),
        ("_xludf.CUSTOM_UDF", "_XLUDF.CUSTOM_UDF"),
    ],
)
def test_xludf_prefix_preserved_for_unknown_addins(raw: str, expected: str) -> None:
    from excel_grapher.core.excel_function_names import (
        normalize_excel_function_name as core_normalize,
    )

    assert core_normalize(raw) == expected
    assert normalize_excel_function_name(raw) == expected


@pytest.mark.parametrize(
    ("canonical", "python_name"),
    [
        ("SUM", "xl_sum"),
        ("IFNA", "xl_ifna"),
        ("XLOOKUP", "xl_xlookup"),
        ("NUMBERVALUE", "xl_numbervalue"),
        ("IFS", "xl_ifs"),
        ("SUMPRODUCT", "xl_sumproduct"),
        ("NORM.DIST", "xl_norm_dist"),
    ],
)
def test_runtime_python_name_follows_xl_prefix_convention(canonical: str, python_name: str) -> None:
    assert excel_func_to_python_runtime_name(canonical) == python_name


@pytest.mark.parametrize(
    "spelling",
    [
        "SUM",
        "_xlfn.SUM",
        "_XLFN.SUM",
        "IFNA",
        "_xlfn.IFNA",
        "XLOOKUP",
        "_xlfn.XLOOKUP",
        "NUMBERVALUE",
        "_xlfn.NUMBERVALUE",
    ],
)
def test_prefixed_spellings_map_to_same_python_runtime_name(spelling: str) -> None:
    bare = spelling.split(".")[-1] if "." in spelling else spelling
    assert excel_func_to_python(spelling) == excel_func_to_python(bare)


def test_functions_registry_has_no_xlfn_alias_keys() -> None:
    """Evaluator dispatch relies on normalization, not per-function ``_XLFN`` keys."""
    xlfn_keys = [key for key in FUNCTIONS if key.startswith("_XLFN.")]
    assert xlfn_keys == []


@pytest.mark.parametrize(
    "pattern",
    [
        'register("_XLFN.',
        "xl__xlfn_",
        "xl__xludf_",
    ],
)
def test_source_tree_has_no_legacy_prefix_handling(pattern: str) -> None:
    """Regression gate: prefix handling stays centralized in ``name_utils``."""
    offenders: list[str] = []
    for path in _SOURCE_ROOT.rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        if pattern in text:
            offenders.append(str(path.relative_to(_REPO_ROOT)))
    assert offenders == []
