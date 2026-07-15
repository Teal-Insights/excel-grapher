"""NumPy is an optional accelerator (`fast` extra), not a required install."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

_CORE = Path(__file__).resolve().parents[3] / "excel_grapher" / "core"
_RUNTIME = Path(__file__).resolve().parents[3] / "excel_grapher" / "runtime"

# Modules that must load (and stay correct) without a top-level NumPy import.
_NUMPY_FREE_TOP_LEVEL = (
    _CORE / "coercions.py",
    _CORE / "math_funcs.py",
    _CORE / "operators_fastpath_stub.py",
    _CORE / "operators_reference.py",
    _CORE / "operator_thresholds.py",
    _RUNTIME / "info.py",
)


def _top_level_numpy_imports(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    found: list[str] = []
    for node in tree.body:
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "numpy" or alias.name.startswith("numpy."):
                    found.append(alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module is not None:
            if node.module == "numpy" or node.module.startswith("numpy."):
                found.append(node.module)
    return found


@pytest.mark.parametrize("path", _NUMPY_FREE_TOP_LEVEL, ids=lambda p: p.name)
def test_library_modules_have_no_top_level_numpy_import(path: Path) -> None:
    assert path.is_file(), f"missing {path}"
    assert _top_level_numpy_imports(path) == []


def test_operators_fastpath_stub_is_noop_and_numpy_free() -> None:
    """Stub always falls back; AST guard above ensures it has no NumPy import."""
    from excel_grapher.core.operators_fastpath_stub import (
        MIN_OPERATOR_FASTPATH_CELLS,
        try_fastpath_arithmetic_array,
        try_fastpath_compare_array,
        try_fastpath_concat_array,
        try_fastpath_sumproduct,
    )

    assert MIN_OPERATOR_FASTPATH_CELLS == 64
    assert try_fastpath_arithmetic_array("+", object(), object()) is None
    assert try_fastpath_compare_array("=", object(), object()) is None
    assert try_fastpath_concat_array(object(), object()) is None
    assert try_fastpath_sumproduct([]) is None


def test_coercions_flatten_and_as_scalar_accept_ndarray_like_without_numpy() -> None:
    """Duck-type ndarray buffers so coercions stay NumPy-free at import time."""
    from excel_grapher.core.coercions import as_scalar, flatten
    from excel_grapher.core.types import XlError

    class FakeArray:
        ndim = 2

        @property
        def flat(self) -> list[int]:
            return [1, 2, 3, 4]

        def tolist(self) -> list[list[int]]:
            return [[1, 2], [3, 4]]

    assert as_scalar(FakeArray()) is XlError.VALUE
    assert list(flatten(FakeArray())) == [1, 2, 3, 4]


def test_has_numpy_flag_matches_import() -> None:
    from excel_grapher.core.numpy_support import HAS_NUMPY

    try:
        import numpy  # noqa: F401
    except ImportError:
        assert HAS_NUMPY is False
    else:
        assert HAS_NUMPY is True
