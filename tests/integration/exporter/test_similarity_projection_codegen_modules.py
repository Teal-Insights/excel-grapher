"""Integration: similarity-compression projected packages import and run."""

from __future__ import annotations

import importlib
import sys
from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter import CodeGenerator, SimilarityAwareCompression
from excel_grapher.grapher.similarity_compression import MockEmbeddingProvider
from tests.fixtures.tiny_dsa.workbook import build_tiny_dsa_workbook

_PACKAGE = "similarity_projected_export_pkg"


def _clear_package_modules() -> None:
    for name in list(sys.modules):
        if name == _PACKAGE or name.startswith(f"{_PACKAGE}."):
            sys.modules.pop(name, None)


def test_similarity_projected_generate_modules_package_runs_and_matches_evaluator(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "tiny_dsa.xlsx"
    build_tiny_dsa_workbook(workbook_path)

    targets = ["Engine!C20", "Engine!H20"]
    graph = create_dependency_graph(
        workbook_path,
        targets,
        load_values=True,
        capture_dependency_provenance=True,
    )

    projection = SimilarityAwareCompression(provider=MockEmbeddingProvider()).project(graph)
    projected = projection.projected_graph
    assert "Engine!C14" not in projected
    assert "Engine!H16" not in projected

    files = CodeGenerator(projection).generate_modules(targets)
    assert "xl_eval" in files["internals.py"]
    assert "def compute_all(" in files["api.py"]

    pkg_dir = tmp_path / _PACKAGE
    for filename, content in files.items():
        pkg_dir.mkdir(parents=True, exist_ok=True)
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module(_PACKAGE)
        generated_targets = pkg.compute_all()
        with FormulaEvaluator(projected) as ev:
            evaluator_results = ev.evaluate(targets)
        assert generated_targets == evaluator_results
    finally:
        sys.path.remove(str(tmp_path))
        _clear_package_modules()
