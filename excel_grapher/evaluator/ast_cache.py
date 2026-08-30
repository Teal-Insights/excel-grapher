"""Bounded LRU cache for parsed formula ASTs."""

from __future__ import annotations

from collections import OrderedDict
from collections.abc import Callable, Mapping
from dataclasses import dataclass

from excel_grapher.evaluator.parser import AstNode

DEFAULT_AST_CACHE_MAXSIZE = 4096

_ParseFn = Callable[[str], AstNode]


@dataclass(frozen=True, slots=True)
class AstCacheInfo:
    """Statistics for an `AstCache` instance."""

    hits: int
    misses: int
    maxsize: int
    currsize: int


class AstCache:
    """LRU cache mapping absolute A1 formula strings to absolute-bound ASTs.

    Keys are stripped `normalized_formula` text (`FormulaStyle.A1_ABSOLUTE`).
    That spelling is a lossy key once trees carry per-axis `RelativeAxis`
    intent, so values must be fully `bind_axes`-resolved (`AbsoluteAxis` only).
    Never store a relative tree under this key: another host that shares the
    spelling would inherit the wrong offsets on the string fallback path.
    """

    def __init__(self, maxsize: int = DEFAULT_AST_CACHE_MAXSIZE) -> None:
        if maxsize < 1:
            raise ValueError("maxsize must be at least 1")
        self._maxsize = maxsize
        self._cache: OrderedDict[str, AstNode] = OrderedDict()
        self._hits = 0
        self._misses = 0

    @property
    def maxsize(self) -> int:
        return self._maxsize

    def __len__(self) -> int:
        return len(self._cache)

    def get(self, normalized_formula: str, *, parse_fn: _ParseFn) -> AstNode:
        """Return a cached AST or parse and store `normalized_formula`."""
        if normalized_formula in self._cache:
            self._hits += 1
            self._cache.move_to_end(normalized_formula)
            return self._cache[normalized_formula]

        self._misses += 1
        ast = parse_fn(normalized_formula)
        self._cache[normalized_formula] = ast
        if len(self._cache) > self._maxsize:
            self._cache.popitem(last=False)
        return ast

    def seed(self, entries: Mapping[str, AstNode]) -> None:
        """Insert pre-parsed absolute-bound ASTs without affecting hit/miss stats.

        Each value must already be host-independent (`bind_axes` against the
        originating `NodeKey`). First-wins: existing keys are left unchanged.
        """
        for normalized_formula, ast in entries.items():
            if normalized_formula in self._cache:
                continue
            self._cache[normalized_formula] = ast
            if len(self._cache) > self._maxsize:
                self._cache.popitem(last=False)

    def clear(self) -> None:
        """Remove all cached ASTs and reset statistics."""
        self._cache.clear()
        self._hits = 0
        self._misses = 0

    def cache_info(self) -> AstCacheInfo:
        """Return hit/miss and size statistics."""
        return AstCacheInfo(
            hits=self._hits,
            misses=self._misses,
            maxsize=self._maxsize,
            currsize=len(self._cache),
        )
