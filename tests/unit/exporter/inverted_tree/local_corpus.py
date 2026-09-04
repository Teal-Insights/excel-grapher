"""Load the local inverted-tree workbook pool and compare export rungs.

See `plans/inverted-tree-scheduling.md` §12 and `tests/fixtures/local/corpus.toml`.
"""

from __future__ import annotations

import inspect
import tomllib
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Literal

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, build_catalog
from excel_grapher.exporter.inverted_tree.deps import SeriesDeps, collect_all_deps
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.schedule import tarjan_series_sccs
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.constraints import (
    constraints_table,
)
from excel_grapher.grapher.constraints import (
    load_constraints_module as load_constraints_module_required,
)
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.load import load_series_bindings
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.paths import FIXTURES_ROOT, LOCAL_CORPUS
from tests.unit.exporter.inverted_tree.helpers import (
    input_kwargs,
)

_MANIFEST = LOCAL_CORPUS / "corpus.toml"


@dataclass(frozen=True, slots=True)
class CorpusEntry:
    """One workbook listed in `corpus.toml`."""

    id: str
    workbook: Path
    bindings: Path
    constraints: Path | None
    max_cells: int
    source: str


def _resolve_entry_path(rel: str, *, source: str) -> Path:
    root = FIXTURES_ROOT if source == "committed" else LOCAL_CORPUS
    return (root / rel).resolve()


def load_corpus_manifest(path: Path | None = None) -> tuple[CorpusEntry, ...]:
    """Parse `corpus.toml` into resolved entries (missing workbooks stay listed)."""
    manifest = path or _MANIFEST
    raw = tomllib.loads(manifest.read_text(encoding="utf-8"))
    entries: list[CorpusEntry] = []
    for item in raw.get("entry", ()):
        source = str(item.get("source", "local"))
        constraints_rel = item.get("constraints")
        entries.append(
            CorpusEntry(
                id=str(item["id"]),
                workbook=_resolve_entry_path(str(item["workbook"]), source=source),
                bindings=_resolve_entry_path(str(item["bindings"]), source=source),
                constraints=(
                    None
                    if constraints_rel is None
                    else _resolve_entry_path(str(constraints_rel), source=source)
                ),
                max_cells=int(item.get("max_cells", 0)),
                source=source,
            )
        )
    return tuple(entries)


def require_workbook(entry: CorpusEntry) -> None:
    """Skip when the pool workbook is not on disk (run-if-available)."""
    if not entry.workbook.is_file():
        pytest.skip(f"local corpus workbook missing: {entry.workbook}")
    if not entry.bindings.exists():
        pytest.skip(f"local corpus bindings missing: {entry.bindings}")


def load_constraints_module(path: Path | None) -> ModuleType | None:
    """Import a constraints module from `path`, or return None."""
    if path is None or not path.is_file():
        return None
    return load_constraints_module_required(path)


def _dynamic_refs(constraints: ModuleType | None) -> DynamicRefConfig | None:
    if constraints is None:
        return None
    table = constraints_table(constraints)
    if not table:
        return None
    return DynamicRefConfig.from_constraints(table, {})


def build_corpus_graph(
    entry: CorpusEntry,
) -> tuple[SeriesCatalog, dict[str, SeriesDeps], DependencyGraph]:
    """Build catalog, deps, and graph for one corpus entry."""
    bindings = load_series_bindings(entry.bindings)
    targets = all_series_targets(bindings, workbook=entry.workbook)
    constraints = load_constraints_module(entry.constraints)
    graph = create_dependency_graph(
        entry.workbook,
        targets,
        load_values=True,
        use_cached_dynamic_refs=constraints is None,
        dynamic_refs=_dynamic_refs(constraints),
        capture_dependency_provenance=True,
    )
    catalog = build_catalog(bindings, workbook=entry.workbook, graph=graph)
    return catalog, collect_all_deps(catalog, graph), graph


def generate_corpus_modules(
    entry: CorpusEntry,
    graph: DependencyGraph,
    _catalog: SeriesCatalog,
    *,
    force_rung: Literal[2, 3] | None = None,
) -> dict[str, str]:
    """Emit inverted-tree modules for a corpus workbook."""
    bindings = load_series_bindings(entry.bindings)
    return generate_inverted_tree_modules(
        graph,
        series_bindings=bindings,
        bindings_workbook=entry.workbook,
        force_rung=force_rung,
    )


def statement_topo_order(catalog: SeriesCatalog, deps: Mapping[str, SeriesDeps]) -> tuple[str, ...]:
    """Return formula series ids in statement-graph topological order."""
    ids = [series.series_id for series in catalog.formula_series()]
    ordered: list[str] = []
    for scc in tarjan_series_sccs(ids, deps):
        ordered.extend(scc)
    return tuple(ordered)


def _values_close(got: object, expected: object) -> bool:
    if isinstance(got, tuple) and isinstance(expected, tuple):
        return len(got) == len(expected) and all(
            _values_close(left, right) for left, right in zip(got, expected, strict=True)
        )
    if isinstance(got, str) or isinstance(expected, str):
        return got == expected
    try:
        return got == pytest.approx(expected)
    except TypeError:
        return got == expected


def _output_cells(series: BoundSeries) -> tuple[str, ...]:
    if series.layout != "matrix" or not series.key_fields:
        return series.cells
    keyed = [
        (tuple(point[field] for field in series.key_fields), cell)
        for point, cell in zip(series.domain, series.cells, strict=True)
    ]
    try:
        keyed.sort(key=lambda item: item[0])
    except TypeError:
        return series.cells
    return tuple(cell for _key, cell in keyed)


def compare_package_to_evaluator(
    pkg: object,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    topo: Sequence[str],
) -> list[str]:
    """Return divergence lines in statement-graph order (root series first)."""
    kwargs = input_kwargs(catalog, graph)
    for series in catalog.constant_series():
        values = []
        for cell in series.cells:
            node = graph.get_node(cell)
            values.append(None if node is None else node.value)
        kwargs[series.series_id] = values[0] if series.is_scalar else tuple(values)
    cells = [cell for series in catalog.formula_series() for cell in series.cells]
    expected = FormulaEvaluator(graph).evaluate(cells)
    lines: list[str] = []
    by_id = {series.series_id: series for series in catalog.formula_series()}
    internals = getattr(pkg, "internals", None)
    for series_id in topo:
        series = by_id[series_id]
        name = series.compute_name or f"compute_{series.series_id}"
        function = getattr(pkg, name, None)
        if function is None and internals is not None:
            function = getattr(internals, series.series_id, None)
        if function is None or not callable(function):
            continue
        accepted = set(inspect.signature(function).parameters)
        got = function(**{key: value for key, value in kwargs.items() if key in accepted})
        kwargs[series.series_id] = got
        if not isinstance(got, tuple):
            got = (got,)
        want = tuple(expected[cell] for cell in _output_cells(series))
        if not _values_close(got, want):
            lines.append(f"{series_id}: export={got!r} evaluator={want!r}")
    return lines


def package_byte_size(modules: Mapping[str, str]) -> int:
    """Return generated package size in bytes (api + internals + runtime)."""
    return sum(
        len(modules[name].encode("utf-8")) for name in ("api.py", "internals.py", "runtime.py")
    )


def assert_no_per_cell_unroll(modules: Mapping[str, str]) -> None:
    """Fail when generated internals embed a per-cell helper surface."""
    internals = modules["internals.py"]
    assert "def cell_" not in internals
    assert "def make_context" not in modules["api.py"]
