"""Codegen embeds vectorized operators only when formulas use large arrays."""

from __future__ import annotations

from pathlib import Path

from excel_grapher import create_dependency_graph
from excel_grapher.exporter.codegen import CodeGenerator
from tests.integration.utils.parity_harness import count_embedded_runtime_lines
from tests.unit.exporter.test_dep_tracking_codegen_emission import _minimal_non_iterative_code
from tests.unit.gaps.workbook_helpers import write_large_string_criteria_sumproduct


def test_minimal_export_omits_vectorized_operator_fastpath() -> None:
    code = _minimal_non_iterative_code()
    assert "batch_coerce_to_float64" not in code


def test_large_sumproduct_export_includes_vectorized_operator_fastpath(
    tmp_path: Path,
) -> None:
    workbook = write_large_string_criteria_sumproduct(
        tmp_path / "large_sumproduct_embed.xlsx",
        rows=2_000,
    )
    graph = create_dependency_graph(
        workbook,
        ["Data!C1"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    code = CodeGenerator(graph).generate(["Data!C1"])
    assert "batch_coerce_to_float64" in code
    assert count_embedded_runtime_lines(code) > count_embedded_runtime_lines(
        _minimal_non_iterative_code()
    )
