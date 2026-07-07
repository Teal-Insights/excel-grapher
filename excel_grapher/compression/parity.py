"""Round-trip value parity between original and expanded AST maps."""

from __future__ import annotations

from collections.abc import Mapping

from excel_grapher.core.formula_ast import AstNode
from excel_grapher.core.types import CellValue

from .types import CompressedNode


def assert_compression_parity(
    original: Mapping[str, AstNode],
    compressed: Mapping[str, CompressedNode],
    *,
    input_values: Mapping[str, CellValue],
    rtol: float = 1e-9,
    atol: float = 0.0,
) -> None:
    """Assert expanded compressed ASTs evaluate to the same values as originals.

    Args:
        original: Per-cell AST map before compression.
        compressed: Mixed compressed map to expand and evaluate.
        input_values: Leaf cell values shared by both evaluation graphs.
        rtol: Relative tolerance for finite float comparison.
        atol: Absolute tolerance for finite float comparison.

    Raises:
        AssertionError: When any target cell value differs.
        NotImplementedError: Until the parity harness is implemented.
    """
    raise NotImplementedError
