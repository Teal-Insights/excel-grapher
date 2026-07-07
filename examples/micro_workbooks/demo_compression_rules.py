#!/usr/bin/env python3
"""Demonstrate formula AST compression rules and pipeline order.

Builds (or reuses) ``compression_rules.xlsx``, extracts per-cell formula ASTs,
then walks the ``compress_full`` pipeline from the compression design doc.
Implemented rules run live; planned rules show the expected artifact shape.

Run from the repo root::

    uv run python examples/micro_workbooks/build_compression_rules_workbook.py
    uv run python examples/micro_workbooks/demo_compression_rules.py
    uv run python examples/micro_workbooks/demo_compression_rules.py --check

Related: ``demo_taco_index.py`` drills into TACO index APIs on ``taco_patterns.xlsx``.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TypeGuard

from excel_grapher import create_dependency_graph
from excel_grapher.compression import (
    COMPRESSION_RULES,
    ParallelFormulaNode,
    SubexpressionRefNode,
    TacoPatternNode,
    apply_compression_rules,
    apply_constant_folding,
    apply_pass_through,
    assert_compression_parity,
    compression_rule_ids,
    empty_compression_stats,
    expand_compressed_to_cells,
    get_rule_apply,
)
from excel_grapher.compression.types import CompressedNode
from excel_grapher.core.address_keys import format_cell_key, normalize_key
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    ColumnVarCellRefNode,
    ErrorNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    parse,
)
from excel_grapher.core.types import CellValue
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.range_compression import PatternKind, build_taco_index
from excel_grapher.grapher.range_compression.types import RangeRef

WORKBOOK = Path(__file__).with_name("compression_rules.xlsx")
BUILD_SCRIPT = Path(__file__).with_name("build_compression_rules_workbook.py")

TARGETS = [
    "Compress!A5:Compress!D5",
    "Compress!D10:Compress!F10",
    "Compress!A14:Compress!C14",
    "Compress!A18:Compress!A20",
    "Taco!D3:D7",
    "Taco!F3:F7",
    "Taco!H3:H7",
    "Taco!K3:K7",
    "Taco!P3:P7",
]

RULE1_CELLS = ("Compress!A5", "Compress!C5", "Compress!D5")
RULE2_CELLS = ("Compress!D10", "Compress!E10", "Compress!F10")
RULE3_CELLS = ("Compress!A14", "Compress!B14", "Compress!C14")
RULE4_CELLS = ("Compress!A18", "Compress!A19", "Compress!A20")


def _is_cell_ast(node: CompressedNode) -> TypeGuard[AstNode]:
    """Return True when `node` is a per-cell formula AST, not a compressed artifact."""
    return not isinstance(node, (ParallelFormulaNode, TacoPatternNode))


def _as_ast_map(cell_map: Mapping[str, CompressedNode]) -> dict[str, AstNode]:
    """Narrow a compressed map to per-cell AST entries."""
    return {key: node for key, node in cell_map.items() if _is_cell_ast(node)}


@dataclass(frozen=True, slots=True)
class PipelineStep:
    """One stage in the ``compress_full`` pipeline."""

    order: str
    rule_ids: tuple[str, ...]
    title: str
    detail: str
    demo_cells: tuple[str, ...] = ()
    apply: (
        Callable[[Mapping[str, AstNode], CompressionStatsHolder], dict[str, CompressedNode]] | None
    ) = None
    illustrate: Callable[..., None] | None = None


@dataclass
class CompressionStatsHolder:
    """Thin wrapper so pipeline steps can share stats."""

    stats: object


def _ensure_workbook() -> None:
    if WORKBOOK.is_file():
        return
    print(f"workbook missing; running {BUILD_SCRIPT.name}")
    subprocess.run(
        [sys.executable, str(BUILD_SCRIPT)],
        check=True,
    )


def ast_map_from_graph(graph: DependencyGraph) -> dict[str, AstNode]:
    """Build a per-cell AST map from an extracted dependency graph."""
    result: dict[str, AstNode] = {}
    for key, node in graph.formula_nodes():
        normalized_formula = node.normalized_formula
        if not normalized_formula:
            continue
        result[normalize_key(key)] = parse(normalized_formula.strip())
    return result


def leaf_values_from_graph(graph: DependencyGraph) -> dict[str, CellValue]:
    """Collect leaf cell values from the graph for parity evaluation."""
    values: dict[str, CellValue] = {}
    for key in graph:
        node = graph.get_node(key)
        if node is None or not node.is_leaf:
            continue
        if node.value is None:
            continue
        values[normalize_key(key)] = node.value
    return values


def ast_to_display(node: AstNode | CompressedNode) -> str:
    """Render an AST node as a formula-like string for console output."""
    if isinstance(node, ParallelFormulaNode):
        cols = f"{node.start_col}{node.output_row}:{node.end_col}{node.output_row}"
        return f"ParallelFormulaNode({node.sheet}!{cols}, template={ast_to_display(node.template)})"
    if isinstance(node, SubexpressionRefNode):
        return node.ref_key
    if isinstance(node, ColumnVarCellRefNode):
        if node.sheet and node.row is not None:
            return f"{node.sheet}!{node.column_variable}{node.row}"
        return node.column_variable
    if isinstance(node, NumberNode):
        value = node.value
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(node, CellRefNode):
        return node.address
    if isinstance(node, RangeNode):
        return f"{node.start}:{node.end}"
    if isinstance(node, UnaryOpNode):
        return f"{node.op}{ast_to_display(node.operand)}"
    if isinstance(node, BinaryOpNode):
        left = ast_to_display(node.left)
        right = ast_to_display(node.right)
        if node.op in {"&"}:
            return f"{left}{node.op}{right}"
        return f"({left}{node.op}{right})"
    if isinstance(node, FunctionCallNode):
        args = ",".join(ast_to_display(arg) for arg in node.args)
        return f"{node.name.upper()}({args})"
    if isinstance(node, StringNode):
        return f'"{node.value}"'
    if isinstance(node, BoolNode):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, ErrorNode):
        return str(node.error)
    return type(node).__name__


def print_cell_map(
    cell_map: Mapping[str, CompressedNode],
    *,
    cells: Sequence[str] | None = None,
    indent: str = "  ",
) -> None:
    """Print selected entries from a compressed node map."""
    keys = [normalize_key(key) for key in cells] if cells is not None else sorted(cell_map)
    for key in keys:
        node = cell_map.get(key)
        if node is None:
            print(f"{indent}{key}: (removed — absorbed into artifact)")
            continue
        print(f"{indent}{key} = {ast_to_display(node)}")


def print_stats(rule_id: str, stats_holder: CompressionStatsHolder) -> None:
    from excel_grapher.compression.stats import CompressionStats

    stats = stats_holder.stats
    if not isinstance(stats, CompressionStats):
        return
    for contribution in stats.rule_contributions:
        if contribution.rule_id != rule_id:
            continue
        parts = []
        if contribution.in_place_transforms:
            parts.append(f"transforms={contribution.in_place_transforms}")
        if contribution.cells_affected:
            parts.append(f"cells_affected={contribution.cells_affected}")
        if contribution.binding_sites:
            parts.append(f"binding_sites={contribution.binding_sites}")
        if contribution.ast_subnodes_saved:
            parts.append(f"ast_subnodes_saved={contribution.ast_subnodes_saved}")
        if parts:
            print(f"  stats: {', '.join(parts)}")
        return


def _apply_pass_through_step(
    ast_map: Mapping[str, AstNode],
    stats_holder: CompressionStatsHolder,
) -> dict[str, CompressedNode]:
    from excel_grapher.compression.stats import CompressionStats

    stats = stats_holder.stats
    assert isinstance(stats, CompressionStats)
    return apply_pass_through(ast_map, stats)


def _apply_constant_folding_step(
    ast_map: Mapping[str, AstNode],
    stats_holder: CompressionStatsHolder,
) -> dict[str, CompressedNode]:
    from excel_grapher.compression.stats import CompressionStats

    stats = stats_holder.stats
    assert isinstance(stats, CompressionStats)
    return apply_constant_folding(ast_map, stats)


def _illustrate_parallel_row(ast_map: Mapping[str, AstNode]) -> None:
    print("  before (from workbook):")
    for key in RULE2_CELLS:
        print(f"    {key} = {ast_to_display(ast_map[normalize_key(key)])}")
    sample = ast_map[normalize_key("Compress!D10")]
    assert isinstance(sample, FunctionCallNode)
    condition, if_true, _if_false = sample.args
    template = FunctionCallNode(
        "IF",
        [
            condition,
            if_true,
            ColumnVarCellRefNode(column_variable="COL", sheet="Ext", row=87),
        ],
    )
    artifact = ParallelFormulaNode(
        sheet="Compress",
        template=template,
        start_col="D",
        end_col="F",
        output_row=10,
        condition=template.args[0],
        if_true=template.args[1],
        if_false=template.args[2],
    )
    print("  expected artifact (not yet emitted by the engine):")
    print(f"    {ast_to_display(artifact)}")
    print("  per-cell expansion:")
    expanded = expand_compressed_to_cells({f"parallel:Compress!{artifact.output_row}": artifact})
    for key in RULE2_CELLS:
        print(f"    {normalize_key(key)} = {ast_to_display(expanded[normalize_key(key)])}")


def _illustrate_cell_cse(ast_map: Mapping[str, AstNode]) -> None:
    print("  before (from workbook):")
    for key in RULE4_CELLS:
        print(f"    {key} = {ast_to_display(ast_map[normalize_key(key)])}")
    shared = parse("=Compress!B17+Compress!C17")
    compressed: dict[str, CompressedNode] = {
        "_cse!0": shared,
        normalize_key("Compress!A18"): BinaryOpNode(
            "*",
            SubexpressionRefNode("_cse!0"),
            NumberNode(2.0),
        ),
        normalize_key("Compress!A19"): BinaryOpNode(
            "*",
            SubexpressionRefNode("_cse!0"),
            NumberNode(3.0),
        ),
        normalize_key("Compress!A20"): BinaryOpNode(
            "+",
            SubexpressionRefNode("_cse!0"),
            NumberNode(10.0),
        ),
    }
    print("  expected compressed map (cell CSE — planned #359):")
    print(f"    _cse!0 = {ast_to_display(shared)}")
    for key in RULE4_CELLS:
        print(f"    {key} = {ast_to_display(compressed[normalize_key(key)])}")
    expanded = expand_compressed_to_cells(compressed)
    print("  after expand (equivalent to originals):")
    for key in RULE4_CELLS:
        print(f"    {key} = {ast_to_display(expanded[normalize_key(key)])}")


def _illustrate_taco(graph: DependencyGraph) -> None:
    index = build_taco_index(graph)
    print("  TACO index on this workbook (grapher.range_compression):")
    kind_order = {
        PatternKind.rr: 0,
        PatternKind.rf: 1,
        PatternKind.fr: 2,
        PatternKind.ff: 3,
        PatternKind.rr_chain: 4,
    }
    for edge in sorted(index.compressed_edges, key=lambda item: kind_order.get(item.meta.kind, 99)):
        dep = _range_label(edge.dependent)
        prec = _range_label(edge.precedent)
        print(f"    {edge.meta.kind}: {dep} <- {prec}")


def _illustrate_artifact_cse(ast_map: Mapping[str, AstNode]) -> None:
    sample = ast_map[normalize_key("Compress!D10")]
    assert isinstance(sample, FunctionCallNode)
    shared = sample.args[1]
    template_a = sample
    template_b = FunctionCallNode(
        "IF",
        [
            BinaryOpNode("=", CellRefNode("Ext!D3"), StringNode("Maybe")),
            shared,
            ColumnVarCellRefNode(column_variable="COL", sheet="Ext", row=88),
        ],
    )
    print("  artifact CSE (planned #360) deduplicates shared template subtrees:")
    print(f"    shared subtree: {ast_to_display(shared)}")
    print(f"    template A tail: {ast_to_display(template_a)}")
    print(f"    template B tail: {ast_to_display(template_b)}")
    print("    → hoist NA() once into _cse!1 referenced from both templates")


def _range_label(ref: RangeRef) -> str:
    if ref.min_col == ref.max_col and ref.min_row == ref.max_row:
        return format_cell_key(ref.sheet, ref.min_col, ref.min_row)
    return (
        f"{format_cell_key(ref.sheet, ref.min_col, ref.min_row)}:"
        f"{format_cell_key(ref.sheet, ref.max_col, ref.max_row)}"
    )


def pipeline_steps() -> tuple[PipelineStep, ...]:
    """Return ``compress_full`` stages in design-doc order."""
    return (
        PipelineStep(
            order="1",
            rule_ids=("pass_through",),
            title="Direct cell reference elimination",
            detail="Replace references to pass-through cells with their ultimate targets.",
            demo_cells=RULE1_CELLS,
            apply=_apply_pass_through_step,
        ),
        PipelineStep(
            order="2",
            rule_ids=("parallel_if_row",),
            title="Parallel row compression",
            detail="Merge contiguous same-row templates into ParallelFormulaNode artifacts.",
            demo_cells=RULE2_CELLS,
        ),
        PipelineStep(
            order="3",
            rule_ids=("constant_folding",),
            title="Constant folding",
            detail="Pre-compute literal-only subexpressions in place.",
            demo_cells=RULE3_CELLS,
            apply=_apply_constant_folding_step,
        ),
        PipelineStep(
            order="4a",
            rule_ids=("common_subexpression",),
            title="Cell CSE fixpoint",
            detail="Hoist repeated subtrees across per-cell formulas to _cse! bindings.",
            demo_cells=RULE4_CELLS,
        ),
        PipelineStep(
            order="4b",
            rule_ids=("constant_folding",),
            title="Post-CSE constant folding",
            detail="Fold any new literal opportunities after CSE substitution.",
            demo_cells=RULE3_CELLS,
        ),
        PipelineStep(
            order="5–9",
            rule_ids=tuple(
                rule_id for rule_id in compression_rule_ids() if rule_id.startswith("taco_")
            ),
            title="TACO range-pattern compression",
            detail="Compress autofill ranges into TacoPatternNode artifacts (RR … RR-Chain).",
        ),
        PipelineStep(
            order="10",
            rule_ids=(),
            title="Remaining per-cell formulas",
            detail="Add formulas not absorbed by parallel or TACO artifacts to the compressed map.",
        ),
        PipelineStep(
            order="4c",
            rule_ids=("common_subexpression",),
            title="Artifact CSE",
            detail="Deduplicate shared subtrees across parallel and TACO templates.",
        ),
    )


def print_pipeline_overview() -> None:
    print("compress_full pipeline order")
    print("=" * 72)
    for step in pipeline_steps():
        ids = ", ".join(step.rule_ids) if step.rule_ids else "(n/a)"
        status = "implemented" if step.apply is not None else "planned / illustrative"
        if step.order == "5–9":
            status = "implemented in grapher.range_compression (compression wiring #360+)"
        print(f"  {step.order:>3}  [{ids}]  {step.title}  ({status})")
    print()
    print("COMPRESSION_RULES metadata order (excel_grapher.compression.rules):")
    for rule in COMPRESSION_RULES:
        wired = "apply wired" if get_rule_apply(rule.rule_id) else "no applier yet"
        emit = "reduces emission units" if rule.reduces_emission_units else "in-place"
        print(f"  {rule.rule_id}: {rule.name} ({wired}; {emit})")
    print()


def run_demo(*, check: bool) -> int:
    _ensure_workbook()
    graph = create_dependency_graph(WORKBOOK, TARGETS, load_values=True)
    original = ast_map_from_graph(graph)
    input_values = leaf_values_from_graph(graph)
    stats_holder = CompressionStatsHolder(stats=empty_compression_stats())

    print(f"workbook: {WORKBOOK.name}")
    print(f"formula cells extracted: {len(original)}")
    print()

    print_pipeline_overview()

    working: dict[str, CompressedNode] = dict(original)
    exit_code = 0

    for step in pipeline_steps():
        print("-" * 72)
        print(f"Step {step.order}: {step.title}")
        print(f"  rule id(s): {', '.join(step.rule_ids) or 'n/a'}")
        print(f"  {step.detail}")

        if step.apply is not None:
            before = {key: working[key] for key in step.demo_cells if key in working}
            if before:
                print("  before:")
                print_cell_map(before, indent="    ")
            ast_only = _as_ast_map(working)
            working = dict(step.apply(ast_only, stats_holder))
            after = {key: working[key] for key in step.demo_cells if key in working}
            if after:
                print("  after:")
                print_cell_map(after, indent="    ")
            for rule_id in step.rule_ids:
                print_stats(rule_id, stats_holder)
            if check and before:
                try:
                    assert_compression_parity(
                        _as_ast_map(before),
                        {normalize_key(k): working[normalize_key(k)] for k in before},
                        input_values=input_values,
                    )
                    print("  parity: ok")
                except AssertionError as exc:
                    print(f"  parity: FAILED\n    {exc}")
                    exit_code = 1
        elif step.order == "5–9":
            _illustrate_taco(graph)
        elif step.order == "2":
            _illustrate_parallel_row(original)
        elif step.order == "4a":
            _illustrate_cell_cse(original)
        elif step.order == "4c":
            _illustrate_artifact_cse(original)
        elif step.illustrate is not None:
            step.illustrate()
        else:
            print("  (runs after upstream rules land — no-op in this demo)")

        print()

    print("-" * 72)
    print("Implemented rules via apply_compression_rules() today:")
    compressed = apply_compression_rules(original)
    implemented_ids = [
        rule_id for rule_id in compression_rule_ids() if get_rule_apply(rule_id) is not None
    ]
    print(f"  active rule ids: {', '.join(implemented_ids)}")
    print("  default pipeline order: pass_through → parallel_if_row → constant_folding")
    print()
    if check:
        try:
            assert_compression_parity(original, compressed, input_values=input_values)
            print("full implemented pipeline parity: ok")
        except AssertionError as exc:
            print(f"full implemented pipeline parity: FAILED\n  {exc}")
            exit_code = 1
    return exit_code


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="Run parity checks for implemented compression steps",
    )
    args = parser.parse_args()
    raise SystemExit(run_demo(check=args.check))


if __name__ == "__main__":
    main()
