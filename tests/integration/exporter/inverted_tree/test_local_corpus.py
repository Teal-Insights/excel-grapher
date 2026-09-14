"""Local workbook pool gate: rung 3 == evaluator, auto == rung 3 (#656).

Opt in with `pytest -m local_corpus`. Missing pool workbooks `pytest.skip`
with the path (parity.mdc run-if-available).
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.paths import LOCAL_CORPUS
from tests.unit.exporter.inverted_tree.helpers import load_package
from tests.unit.exporter.inverted_tree.local_corpus import (
    CorpusEntry,
    assert_no_per_cell_unroll,
    build_corpus_graph,
    compare_package_to_evaluator,
    generate_corpus_modules,
    load_corpus_manifest,
    package_byte_size,
    require_workbook,
    statement_topo_order,
)


@pytest.fixture(scope="module")
def corpus_entries() -> tuple[CorpusEntry, ...]:
    return load_corpus_manifest()


def test_local_corpus_path_and_manifest() -> None:
    assert LOCAL_CORPUS.name == "local"
    entries = load_corpus_manifest()
    ids = [entry.id for entry in entries]
    assert ids == ["tiny_dsa", "qcraft", "lic_dsf"]
    tiny = entries[0]
    assert tiny.workbook.is_file()
    assert tiny.source == "committed"


@pytest.mark.local_corpus
@pytest.mark.parametrize("entry_id", ["tiny_dsa", "qcraft", "lic_dsf"])
def test_corpus_rung3_matches_evaluator_and_auto(
    tmp_path: Path,
    corpus_entries: tuple[CorpusEntry, ...],
    entry_id: str,
) -> None:
    entry = next(item for item in corpus_entries if item.id == entry_id)
    require_workbook(entry)
    catalog, deps, graph = build_corpus_graph(entry)
    if entry.max_cells and sum(len(series.cells) for series in catalog.formula_series()) > (
        entry.max_cells
    ):
        pytest.skip(f"{entry.id} exceeds max_cells={entry.max_cells}")
    topo = statement_topo_order(catalog, deps)
    modules = generate_corpus_modules(entry, graph, catalog)
    pkg = load_package(modules, tmp_path, name=f"{entry.id}_auto")
    lines = compare_package_to_evaluator(pkg, catalog, graph, topo=topo)
    assert not lines, "export diverged from evaluator:\n" + "\n".join(lines)
    assert_no_per_cell_unroll(modules)
    assert package_byte_size(modules) == package_byte_size(
        generate_corpus_modules(entry, graph, catalog)
    )
