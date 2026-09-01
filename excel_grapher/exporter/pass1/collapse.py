"""Collapse bound cell_* IR into Pass-1 series helpers (issue #595).

Runs after CodeGenerator emits per-cell translations. Bound internal/output
formula series become def <series_id>(ctx, ...) helpers; unbound leftovers may
remain as cell_*. Verification mismatch raises SeriesHelperVerificationError.
"""

from __future__ import annotations

import ast
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_grapher.core.address_keys import normalize_key as normalize_address
from excel_grapher.evaluator.name_utils import address_to_python_name
from excel_grapher.exporter.pass1.addresses import parse_workbook_address
from excel_grapher.exporter.pass1.bindings import (
    BindingKeyValue,
    KeyConceptSpec,
    build_address_to_series_id,
    build_bound_address_keys,
    expected_member_keys_for_cluster,
    helper_parameters_for_varying_keys,
    key_concept_vocabulary_from_bindings,
    render_literal_helper_call,
    varying_key_concepts,
)
from excel_grapher.exporter.pass1.clustering import FormulaCluster, cluster_graph_formulas
from excel_grapher.exporter.pass1.fingerprints import (
    SemanticDependencyRef,
    build_cluster_fingerprint_summary,
)
from excel_grapher.exporter.pass1.mechanical_body import (
    MechanicalSynthesisError,
    synthesize_cluster_body,
    synthesize_singleton_body,
)
from excel_grapher.exporter.pass1.models import MemberContext, SeriesHelperVerificationError
from excel_grapher.exporter.pass1.modes import ClusteringMode
from excel_grapher.exporter.pass1.naming import allocate_schedule_helper_names
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.constant_series import derive_constant_series
from excel_grapher.series_bindings.input_series import derive_input_series
from excel_grapher.series_bindings.internal_series import derive_internal_series
from excel_grapher.series_bindings.output_series import derive_output_series
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

MECHANICAL_PLACEHOLDER_DOCSTRING = '"""Mechanically synthesized helper pending semantic naming."""'

_XL_EVAL_CELL_CALL = re.compile(
    r"xl_eval\(\s*ctx\s*,\s*(['\"])(?P<addr>[^'\"]+)\1\s*,\s*(?P<fn>cell_[A-Za-z0-9_]+)\s*\)"
)
_CELL_CALL = re.compile(r"\b(?P<fn>cell_[A-Za-z0-9_]+)\s*\(\s*ctx\s*\)")


@dataclass(frozen=True)
class Pass1CollapseResult:
    """Outcome of collapsing bound series helpers into internals IR."""

    source: str
    helper_names: tuple[str, ...]
    address_helpers: dict[str, dict[str, Any]]
    address_dispatch: dict[str, tuple[str, dict[str, BindingKeyValue]]]


def _dtype_annotation(dtype: str) -> str:
    mapping = {
        "int": "int",
        "float": "float",
        "bool": "bool",
        "string": "str",
        "datetime": "datetime.datetime",
    }
    return mapping.get(dtype, "object")


def _indent_body(body: str) -> str:
    lines = body.splitlines() or ["pass"]
    return "\n".join(f"    {line}" if line else "" for line in lines)


def _assemble_helper_source(
    helper_name: str,
    body: str,
    *,
    parameters: Sequence[KeyConceptSpec] = (),
    memoize: bool = False,
) -> str:
    if parameters:
        params = ", ".join(
            ["ctx"]
            + [
                f"{spec.suggested_param_name}: {_dtype_annotation(spec.dtype)}"
                for spec in parameters
            ]
        )
    else:
        params = "ctx"
    chunks: list[str] = []
    if memoize:
        chunks.append("@xl_memoize")
    chunks.append(f"def {helper_name}({params}):")
    chunks.append(f"    {MECHANICAL_PLACEHOLDER_DOCSTRING}")
    chunks.append(_indent_body(body.rstrip()))
    chunks.append("")
    return "\n".join(chunks) + "\n"


def _function_sources(source: str) -> dict[str, str]:
    module = ast.parse(source)
    lines = source.splitlines(keepends=True)
    out: dict[str, str] = {}
    for node in module.body:
        if isinstance(node, ast.FunctionDef):
            out[node.name] = "".join(lines[node.lineno - 1 : node.end_lineno])
    return out


def _replace_function_def(source: str, old_name: str, new_source: str) -> str:
    module = ast.parse(source)
    lines = source.splitlines(keepends=True)
    for node in module.body:
        if isinstance(node, ast.FunctionDef) and node.name == old_name:
            start = node.lineno - 1
            end = node.end_lineno or node.lineno
            replacement = new_source if new_source.endswith("\n") else new_source + "\n"
            if not replacement.endswith("\n\n") and end < len(lines):
                replacement = replacement.rstrip("\n") + "\n\n"
            return "".join(lines[:start]) + replacement + "".join(lines[end:])
    raise KeyError(f"Function {old_name!r} not found")


def _remove_function_defs(source: str, names: set[str]) -> str:
    if not names:
        return source
    module = ast.parse(source)
    lines = source.splitlines(keepends=True)
    spans: list[tuple[int, int]] = []
    for node in module.body:
        if isinstance(node, ast.FunctionDef) and node.name in names:
            spans.append((node.lineno - 1, node.end_lineno or node.lineno))
    if not spans:
        return source
    spans.sort()
    parts: list[str] = []
    cursor = 0
    for start, end in spans:
        parts.append("".join(lines[cursor:start]))
        cursor = end
        if cursor < len(lines) and lines[cursor].strip() == "":
            cursor += 1
    parts.append("".join(lines[cursor:]))
    return "".join(parts)


def _rewrite_call_sites(
    source: str,
    *,
    bindings: Sequence[tuple[str, str, str]],
) -> str:
    by_fn = {fn: call for _addr, fn, call in bindings}
    by_addr = {normalize_address(addr): call for addr, _fn, call in bindings}

    def replace_xl_eval(match: re.Match[str]) -> str:
        addr = normalize_address(match.group("addr"))
        fn = match.group("fn")
        if addr in by_addr:
            return by_addr[addr]
        if fn in by_fn:
            return by_fn[fn]
        return match.group(0)

    def replace_cell_call(match: re.Match[str]) -> str:
        fn = match.group("fn")
        if fn in by_fn:
            return by_fn[fn]
        return match.group(0)

    updated = _XL_EVAL_CELL_CALL.sub(replace_xl_eval, source)
    return _CELL_CALL.sub(replace_cell_call, updated)


def _member_context(
    address: str,
    *,
    graph: DependencyGraph,
    function_sources: Mapping[str, str],
    bound_address_keys: Mapping[str, Mapping[str, BindingKeyValue]],
) -> MemberContext | None:
    function_name = address_to_python_name(address)
    python_source = function_sources.get(function_name)
    if python_source is None:
        return None
    node = graph.get_node(address)
    if node is None or not node.has_formula:
        return None
    formula = node.normalized_formula or ""
    _sheet, column, _row = parse_workbook_address(address)
    deps = tuple(sorted(graph.get_dependencies(address)))
    dep_fns = tuple(
        address_to_python_name(dep)
        for dep in deps
        if (dep_node := graph.get_node(dep)) is not None and dep_node.has_formula
    )
    keys = bound_address_keys.get(address)
    return MemberContext(
        address=address,
        function_name=function_name,
        engine_column=column,
        normalized_formula=formula,
        python_source=python_source,
        dependency_addresses=deps,
        dependency_functions=dep_fns,
        binding_keys=dict(keys) if keys else None,
        binding_record=None,
    )


def _topo_clusters(
    clusters: Sequence[FormulaCluster],
    *,
    graph: DependencyGraph,
) -> list[FormulaCluster]:
    member_of: dict[str, int] = {}
    for index, cluster in enumerate(clusters):
        for address in cluster.members:
            member_of[address] = index

    successors: dict[int, set[int]] = {i: set() for i in range(len(clusters))}
    indegree = [0] * len(clusters)
    for index, cluster in enumerate(clusters):
        deps: set[int] = set()
        for address in cluster.members:
            for dep in graph.get_dependencies(address):
                other = member_of.get(dep)
                if other is not None and other != index:
                    deps.add(other)
        for other in deps:
            if index not in successors[other]:
                successors[other].add(index)
                indegree[index] += 1

    ready = [i for i, degree in enumerate(indegree) if degree == 0]
    ordered: list[int] = []
    while ready:
        i = ready.pop(0)
        ordered.append(i)
        for j in sorted(successors[i]):
            indegree[j] -= 1
            if indegree[j] == 0:
                ready.append(j)
    if len(ordered) != len(clusters):
        return list(clusters)
    return [clusters[i] for i in ordered]


def _ensure_xl_memoize_import(source: str) -> str:
    if "xl_memoize" not in source:
        return source
    pattern = re.compile(r"from \.runtime import \((?P<body>[^)]*)\)", re.DOTALL)
    match = pattern.search(source)
    if match and "xl_memoize" not in match.group("body"):
        body = match.group("body").rstrip()
        if body and not body.endswith(","):
            body = body + ","
        insertion = f"{body}\n    xl_memoize,\n"
        return source[: match.start("body")] + insertion + source[match.end("body") :]
    single = re.compile(r"from \.runtime import (?P<names>[^\n]+)")
    sm = single.search(source)
    if sm and "(" not in sm.group("names") and "xl_memoize" not in sm.group("names"):
        names = sm.group("names").rstrip()
        return source[: sm.start("names")] + f"{names}, xl_memoize" + source[sm.end("names") :]
    return source


def _summary_has_self_recurrence(summary: object) -> bool:
    groups = getattr(summary, "groups", ())
    for group in groups:
        for relation in getattr(group, "ref_relations", ()):
            resolution = getattr(relation, "resolution", None)
            if resolution is not None and getattr(resolution, "kind", None) == ("self_recurrence"):
                return True
    return False


def _series_maps_for_graph(
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    *,
    workbook: Path,
) -> tuple[
    Mapping[str, Mapping[str, BindingKeyValue]],
    Mapping[str, str],
]:
    internal_series = derive_internal_series(graph, bindings, workbook=workbook)
    output_series = derive_output_series(graph, bindings, workbook=workbook)
    input_series = derive_input_series(graph, bindings, workbook=workbook)
    constant_series = derive_constant_series(graph, bindings, workbook=workbook)
    bound_address_keys = build_bound_address_keys(
        input_series,
        output_series,
        internal_series,
        constant_series=constant_series,
    )
    address_to_series_id = build_address_to_series_id(
        internal_series,
        output_series=output_series,
        input_series=input_series,
        constant_series=constant_series,
    )
    return bound_address_keys, address_to_series_id


def _bound_formula_addresses_missing_ir(
    source: str,
    *,
    graph: DependencyGraph,
    address_to_series_id: Mapping[str, str],
) -> list[str]:
    functions = _function_sources(source)
    missing: list[str] = []
    for address in address_to_series_id:
        node = graph.get_node(address)
        if node is None or not node.has_formula:
            continue
        if address_to_python_name(address) not in functions:
            missing.append(address)
    return sorted(missing)


def _member_key_kwargs(
    parameters: Sequence[KeyConceptSpec],
    keys: Mapping[str, BindingKeyValue],
) -> dict[str, BindingKeyValue]:
    return {
        spec.suggested_param_name: keys[spec.dimension_id]
        for spec in parameters
        if spec.dimension_id in keys
    }


def collapse_bound_series_in_source(
    source: str,
    *,
    graph: DependencyGraph,
    bindings: WorkbookSeriesBindings,
    workbook: Path | str,
    canonical_graph: DependencyGraph | None = None,
    clustering_mode: ClusteringMode = "series",
) -> Pass1CollapseResult:
    """Collapse bound formula series in internals IR into named helpers.

    ``graph`` must be the graph that ``cell_*`` IR was emitted from (the
    emission / projected graph). ``canonical_graph`` defaults to ``graph`` and
    is used to detect bound formula addresses whose IR is missing — typically
    because OptimalCompression inlined internals that were not in ``preserve``.
    """
    workbook_path = Path(workbook)
    empty = Pass1CollapseResult(
        source=source,
        helper_names=(),
        address_helpers={},
        address_dispatch={},
    )

    inventory_graph = canonical_graph if canonical_graph is not None else graph
    _bound_keys_canon, address_to_series_id_canon = _series_maps_for_graph(
        inventory_graph,
        bindings,
        workbook=workbook_path,
    )
    if not address_to_series_id_canon:
        return empty

    missing_ir = _bound_formula_addresses_missing_ir(
        source,
        graph=inventory_graph,
        address_to_series_id=address_to_series_id_canon,
    )
    if missing_ir:
        preview = ", ".join(missing_ir[:8])
        more = f" (+{len(missing_ir) - 8} more)" if len(missing_ir) > 8 else ""
        raise SeriesHelperVerificationError(
            "Pass-1 needs cell_* IR for every bound formula address, but these are "
            "missing. If you projected with OptimalCompression, preserve all bound "
            f"series (including internals): {preview}{more}"
        )

    bound_address_keys, address_to_series_id = _series_maps_for_graph(
        graph,
        bindings,
        workbook=workbook_path,
    )
    if not address_to_series_id:
        return empty

    vocabulary = key_concept_vocabulary_from_bindings(bindings)
    clusters = cluster_graph_formulas(
        graph,
        bound_address_keys=bound_address_keys,
        clustering_mode=clustering_mode,
        address_to_series_id=address_to_series_id,
        workbook_path=workbook_path,
    )
    collapsible = [
        cluster
        for cluster in clusters
        if cluster.members
        and all(address in address_to_series_id for address in cluster.members)
        and all(
            (node := graph.get_node(address)) is not None and node.has_formula
            for address in cluster.members
        )
    ]
    if not collapsible:
        bound_formula_leftovers = [
            address
            for address, _series_id in address_to_series_id.items()
            if (node := graph.get_node(address)) is not None and node.has_formula
        ]
        if bound_formula_leftovers:
            preview = ", ".join(bound_formula_leftovers[:8])
            raise SeriesHelperVerificationError(
                f"Pass-1 found bound formula addresses with no collapsible cluster: {preview}"
            )
        return empty

    collapsible = _topo_clusters(collapsible, graph=graph)
    helper_names = allocate_schedule_helper_names(
        [cluster.members for cluster in collapsible],
        address_to_series_id,
    )

    updated = source
    emitted_helpers: list[str] = []
    address_helpers: dict[str, dict[str, Any]] = {}
    address_dispatch: dict[str, tuple[str, dict[str, BindingKeyValue]]] = {}
    semantic_deps: list[SemanticDependencyRef] = []
    needs_memoize = False

    for cluster, helper_name in zip(collapsible, helper_names, strict=True):
        function_sources = _function_sources(updated)
        members: list[MemberContext] = []
        for address in cluster.members:
            member = _member_context(
                address,
                graph=graph,
                function_sources=function_sources,
                bound_address_keys=bound_address_keys,
            )
            if member is None:
                raise SeriesHelperVerificationError(
                    f"missing cell_* IR for bound address {address!r} "
                    f"(series helper {helper_name!r})"
                )
            members.append(member)

        try:
            if len(members) == 1:
                draft = synthesize_singleton_body(
                    members[0].python_source,
                    inline_replacements={},
                )
                helper_source = _assemble_helper_source(
                    helper_name,
                    draft.body,
                    parameters=(),
                    memoize=False,
                )
                parameters: tuple[KeyConceptSpec, ...] = ()
                member_keys: dict[str, dict[str, BindingKeyValue]] = {members[0].address: {}}
            else:
                expected_keys = expected_member_keys_for_cluster(
                    [m.address for m in members],
                    bound_address_keys=bound_address_keys,
                )
                summary = build_cluster_fingerprint_summary(
                    members,
                    expected_member_keys=expected_keys,
                    bound_address_keys=bound_address_keys,
                    address_to_series_id=address_to_series_id,
                    semantic_dependencies=tuple(semantic_deps),
                    workbook_path=workbook_path,
                )
                if summary.fallback_reason is not None:
                    raise MechanicalSynthesisError(
                        f"fingerprint_fallback:{summary.fallback_reason}"
                    )
                # Mixed-regime / key-dispatch is out of MVP (#595). Multiple
                # fingerprint groups are allowed only for self-recurrence.
                if len(summary.groups) > 1 and not _summary_has_self_recurrence(summary):
                    skeletons = sorted({g.skeleton_text for g in summary.groups})
                    raise MechanicalSynthesisError("mixed_regime_groups:" + "|".join(skeletons[:4]))
                draft = synthesize_cluster_body(
                    summary,
                    key_vocabulary=vocabulary,
                    expected_member_keys=expected_keys,
                    helper_name=helper_name,
                )
                varying = varying_key_concepts(
                    [m.address for m in members],
                    bound_address_keys=bound_address_keys,
                )
                parameters = helper_parameters_for_varying_keys(varying, vocabulary)
                helper_source = _assemble_helper_source(
                    helper_name,
                    draft.body,
                    parameters=parameters,
                    memoize=True,
                )
                needs_memoize = True
                member_keys = expected_keys
        except MechanicalSynthesisError as exc:
            raise SeriesHelperVerificationError(
                f"Pass-1 verification failed for helper {helper_name!r}: {exc}"
            ) from exc

        primary = members[0].function_name
        updated = _replace_function_def(updated, primary, helper_source)
        remove = {m.function_name for m in members[1:]}
        updated = _remove_function_defs(updated, remove)

        param_pairs = tuple((spec.suggested_param_name, spec.dimension_id) for spec in parameters)
        collapse_bindings: list[tuple[str, str, str]] = []
        for member in members:
            keys = member_keys.get(member.address, {})
            if parameters:
                literal = render_literal_helper_call(helper_name, param_pairs, keys)
            else:
                literal = f"{helper_name}(ctx)"
            collapse_bindings.append((member.address, member.function_name, literal))
            address_helpers[member.address] = {
                "name": helper_name,
                "dims": [spec.dimension_id for spec in parameters],
            }
            address_dispatch[member.address] = (
                helper_name,
                _member_key_kwargs(parameters, keys),
            )

        updated = _rewrite_call_sites(updated, bindings=collapse_bindings)
        emitted_helpers.append(helper_name)

        semantic_deps.append(
            SemanticDependencyRef(
                helper_name=helper_name,
                call_form=(
                    f"{helper_name}(ctx, "
                    + ", ".join(f"{spec.suggested_param_name}=..." for spec in parameters)
                    + ")"
                    if parameters
                    else f"{helper_name}(ctx)"
                ),
                address_template=members[0].address,
                addresses=tuple(m.address for m in members),
            )
        )

    if needs_memoize:
        updated = _ensure_xl_memoize_import(updated)

    # Fail closed: every formula-backed bound address must have been collapsed.
    remaining = _function_sources(updated)
    leftovers = [
        address
        for address, series_id in address_to_series_id.items()
        if (node := graph.get_node(address)) is not None
        and node.has_formula
        and address_to_python_name(address) in remaining
    ]
    if leftovers:
        preview = ", ".join(leftovers[:8])
        more = f" (+{len(leftovers) - 8} more)" if len(leftovers) > 8 else ""
        raise SeriesHelperVerificationError(
            "Pass-1 left bound formula addresses as cell_* leftovers "
            f"(verification incomplete): {preview}{more}"
        )

    return Pass1CollapseResult(
        source=updated,
        helper_names=tuple(emitted_helpers),
        address_helpers=address_helpers,
        address_dispatch=address_dispatch,
    )
