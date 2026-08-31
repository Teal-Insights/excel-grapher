"""Group interned formula shapes into Excel shared formulas for write-back."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterator
from dataclasses import dataclass
from typing import Literal

from fastpyxl.compat import safe_string
from fastpyxl.worksheet.formula import ArrayFormula

from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    FormulaStyle,
    FunctionCallNode,
    UnaryOpNode,
    render_formula,
)
from excel_grapher.core.formula_shape import (
    AddressLeaf,
    FormulaShapeTable,
    specialize_formula_shape,
)

from .graph import GraphReadView
from .node import NodeView

SharedFormulasMode = Literal["auto", "off", "require"]

_SHARED_FORMULA_MODES: tuple[SharedFormulasMode, ...] = ("auto", "off", "require")


class SharedFormula(ArrayFormula):
    """fastpyxl cell value that persists `t="shared"` formula XML.

    Subclasses `ArrayFormula` so fastpyxl's writer and type detection treat
    the value as a formula and copy `__iter__` attributes onto `<f>`.
    """

    t = "shared"

    def __init__(self, si: int, ref: str | None = None, text: str | None = None) -> None:
        super().__init__(ref, text)
        self.si = si

    def __iter__(self) -> Iterator[tuple[str, str]]:
        yield "t", self.t
        yield "si", str(self.si)
        if self.ref:
            yield "ref", safe_string(self.ref)


def parse_shared_formulas_mode(mode: str) -> SharedFormulasMode:
    """Return a validated `shared_formulas` mode.

    Raises:
        ValueError: If `mode` is not `auto`, `off`, or `require`.
    """
    if mode == "auto":
        return "auto"
    if mode == "off":
        return "off"
    if mode == "require":
        return "require"
    allowed = ", ".join(_SHARED_FORMULA_MODES)
    raise ValueError(f"shared_formulas must be one of {allowed}; got {mode!r}")


def shape_table_for_write(graph: GraphReadView) -> FormulaShapeTable | None:
    """Return a warm `FormulaShapeTable` on `graph` if present.

    Reads `graph.formula_shapes`, then `projected_graph.formula_shapes` so a
    `ProjectionResult` can group when the projected clone was rewarmed. Does
    not rewarm.
    """
    table = getattr(graph, "formula_shapes", None)
    if isinstance(table, FormulaShapeTable):
        return table
    projected = getattr(graph, "projected_graph", None)
    if projected is not None:
        table = getattr(projected, "formula_shapes", None)
        if isinstance(table, FormulaShapeTable):
            return table
    return None


def shared_formula_cell_values(
    graph: GraphReadView,
    *,
    style: FormulaStyle,
    coerce_relative_refs: bool,
    mode: SharedFormulasMode,
) -> dict[str, SharedFormula]:
    """Map node keys to `SharedFormula` values for contiguous autofill runs.

    Groups formula cells that share an interned shape id and identical
    remaining address params (relative axis vectors). Contiguous runs in one
    column or row become one shared-formula group. The master cell formula is
    spelled with `style` (Excel/fastpyxl persist A1 on disk). Group identity
    is the interned relative shape, which matches `render_formula(...,
    style=R1C1)` for those params.

    Missing or stale `formula_shapes`, non-contiguous or mixed-axis leftovers,
    array formulas, `INDIRECT`, and `coerce_relative_refs=True` skip grouping
    and emit per-cell formulas. Does not auto-rewarm the overlay.

    Args:
        graph: Read view whose formula cells may be grouped.
        style: Formula spelling for the master cell body.
        coerce_relative_refs: When True, skip grouping (absolute `$` bindings
            would not match a relative shared fill).
        mode: `auto` groups when shapes are warm; `off` never groups;
            `require` fails if the overlay is missing.

    Returns:
        Node-key to `SharedFormula` for cells that participate in a group.

    Raises:
        ValueError: If `mode` is `require` and `formula_shapes` is missing.
    """
    if mode == "off":
        return {}
    table = shape_table_for_write(graph)
    if table is None:
        if mode == "require":
            raise ValueError(
                "shared_formulas='require' needs a warm graph.formula_shapes overlay; "
                "pass warm_formula_shapes=True on extract or call warm_formula_shapes "
                "and assign the result (the writer does not auto-rewarm)"
            )
        return {}
    if coerce_relative_refs:
        return {}

    sites_by_group: dict[tuple[str, str, tuple[AddressLeaf, ...]], list[_FormulaSite]] = (
        defaultdict(list)
    )
    for key in graph:
        node = graph.get_node(key)
        if node is None:
            continue
        site = _eligible_site(key, node, table)
        if site is None:
            continue
        sites_by_group[(site.sheet, site.shape_key, site.params)].append(site)

    assignments: dict[str, SharedFormula] = {}
    next_si: dict[str, int] = defaultdict(int)
    for sites in sites_by_group.values():
        for run in _contiguous_runs(sites):
            si = next_si[run[0].sheet]
            next_si[run[0].sheet] += 1
            ref = f"{run[0].coord}:{run[-1].coord}"
            master_text = _master_text(run[0].node, key=run[0].key, style=style)
            assignments[run[0].key] = SharedFormula(si, ref=ref, text=master_text)
            for sibling in run[1:]:
                assignments[sibling.key] = SharedFormula(si)
    return assignments


@dataclass(frozen=True, slots=True)
class _FormulaSite:
    key: str
    sheet: str
    col_idx: int
    row: int
    coord: str
    node: NodeView
    shape_key: str
    params: tuple[AddressLeaf, ...]


def _eligible_site(
    key: str,
    node: NodeView,
    table: FormulaShapeTable,
) -> _FormulaSite | None:
    if not node.has_formula or node.formula_ast is None or node.is_array_formula:
        return None
    if not node.sheet or not node.column or node.row is None:
        return None
    if _ast_contains_indirect(node.formula_ast):
        return None
    bound = table.lookup(key)
    if bound is None:
        return None
    shape_key, skeleton, params = bound
    try:
        rebuilt = specialize_formula_shape(skeleton, params)
    except ValueError:
        return None
    if rebuilt != node.formula_ast:
        return None
    return _FormulaSite(
        key=key,
        sheet=node.sheet,
        col_idx=node.column_index,
        row=int(node.row),
        coord=f"{node.column}{node.row}",
        node=node,
        shape_key=shape_key,
        params=params,
    )


def _contiguous_runs(sites: list[_FormulaSite]) -> list[list[_FormulaSite]]:
    """Return column-then-row contiguous runs of length >= 2."""
    unused = list(sites)
    runs: list[list[_FormulaSite]] = []
    for run in _axis_runs(unused, axis="col"):
        if len(run) >= 2:
            runs.append(run)
            for site in run:
                unused.remove(site)
    for run in _axis_runs(unused, axis="row"):
        if len(run) >= 2:
            runs.append(run)
    return runs


def _axis_runs(
    sites: list[_FormulaSite], *, axis: Literal["col", "row"]
) -> list[list[_FormulaSite]]:
    grouped: dict[int, list[_FormulaSite]] = defaultdict(list)
    for site in sites:
        grouped[site.col_idx if axis == "col" else site.row].append(site)
    runs: list[list[_FormulaSite]] = []
    for bucket in grouped.values():
        if not bucket:
            continue
        if axis == "col":
            ordered = sorted(bucket, key=lambda site: site.row)
        else:
            ordered = sorted(bucket, key=lambda site: site.col_idx)
        current = [ordered[0]]
        for site in ordered[1:]:
            prev = current[-1]
            adjacent = (
                site.row == prev.row + 1 if axis == "col" else site.col_idx == prev.col_idx + 1
            )
            if adjacent:
                current.append(site)
            else:
                runs.append(current)
                current = [site]
        runs.append(current)
    return runs


def _master_text(node: NodeView, *, key: str, style: FormulaStyle) -> str:
    ast = node.formula_ast
    if ast is None:
        raise ValueError(f"Cannot write unparseable formula at {key}")
    try:
        return render_formula(ast, anchor=node.address, style=style)
    except ValueError as exc:
        label = str(node.address) if node.address is not None else key
        raise ValueError(f"Cannot render formula at {label}: {exc}") from exc


def _ast_contains_indirect(node: AstNode) -> bool:
    match node:
        case FunctionCallNode(name, args):
            if name.upper() == "INDIRECT":
                return True
            return any(_ast_contains_indirect(arg) for arg in args)
        case BinaryOpNode(_, left, right):
            return _ast_contains_indirect(left) or _ast_contains_indirect(right)
        case UnaryOpNode(_, operand):
            return _ast_contains_indirect(operand)
        case _:
            return False
