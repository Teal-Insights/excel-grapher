"""Type aliases for the compression package."""

from __future__ import annotations

from typing import TypeAlias

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.core.formula_ast import AstNode

from .nodes import ParallelFormulaNode, TacoPatternNode, TemplateAstNode

CompressedNode: TypeAlias = AstNode | ParallelFormulaNode | TacoPatternNode

_SYNTHETIC_KEY_PREFIXES = ("parallel:", "taco:", "_cse!")


def is_synthetic_compressed_key(key: str) -> bool:
    """Return True for artifact or binding keys that are not workbook cell addresses."""
    return key.startswith(_SYNTHETIC_KEY_PREFIXES)


def normalize_compressed_key(key: str) -> str:
    """Normalize a compressed-map key, preserving synthetic artifact keys."""
    if is_synthetic_compressed_key(key):
        return key
    return normalize_key(key)


__all__ = [
    "CompressedNode",
    "TemplateAstNode",
    "is_synthetic_compressed_key",
    "normalize_compressed_key",
]
