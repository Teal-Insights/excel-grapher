from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import cast

from excel_grapher.evaluator.codegen import CodeGenerator, GraphLike
from excel_grapher.evaluator.types import XlError


class _WeirdRepr:
    def __repr__(self) -> str:  # pragma: no cover
        # This is intentionally not a Python literal; embedding it would break generated code.
        return "<weird>"


@dataclass
class _Node:
    # Minimal surface area consumed by CodeGenerator.
    formula: str | None
    normalized_formula: str | None
    value: object | None


class _FakeGraph:
    def __init__(self, nodes: dict[str, _Node], deps: dict[str, list[str]] | None = None) -> None:
        self._nodes = nodes
        self._deps = deps or {}

    def get_node(self, address: str) -> _Node | None:  # noqa: D401
        return self._nodes.get(address)

    def leaf_keys(self) -> list[str]:
        return [k for k, n in self._nodes.items() if n.formula is None]

    def formula_keys(self) -> list[str]:
        return [k for k, n in self._nodes.items() if n.formula is not None]

    def dependencies(self, address: str) -> list[str]:
        return self._deps.get(address, [])


def test_codegen_does_not_embed_non_literal_leaf_values() -> None:
    graph = _FakeGraph(
        nodes={
            "S!A1": _Node(formula=None, normalized_formula=None, value=_WeirdRepr()),
        }
    )
    code = CodeGenerator(cast(GraphLike, graph)).generate(["S!A1"])
    compiled = compile(code, "<generated>", "exec")
    assert compiled is not None
    assert "<weird>" not in code


def test_codegen_serializes_xlerror_leaf_values() -> None:
    graph = _FakeGraph(
        nodes={
            "S!A1": _Node(formula=None, normalized_formula=None, value=XlError.DIV),
        }
    )
    code = CodeGenerator(cast(GraphLike, graph)).generate(["S!A1"])
    compiled = compile(code, "<generated>", "exec")
    assert compiled is not None

    namespace: dict[str, object] = {}
    exec(code, namespace)
    compute_all = cast(Callable[[], dict[str, object]], namespace["compute_all"])
    results = compute_all()
    assert results["S!A1"] == XlError.DIV
