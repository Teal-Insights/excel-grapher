"""Column OFFSET from a series with no column axis (#991).

A one-column anchor and the translation block beside it are separate series.
`OFFSET(anchor, 0, language)` steps the column key that covers the in-domain
landings, starting at the anchor column.
"""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from excel_grapher import CodeGenerator, DynamicRefConfig, create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings import load_series_bindings
from tests.unit.exporter.inverted_tree.helpers import invoke_public_compute, load_package

_PHRASES = (
    (
        "Change in PPG debt",
        "Change in PPG debt (fr)",
        "Change in PPG debt (pt)",
        "Change in PPG debt (es)",
    ),
    ("Median", "Median (fr)", "Median (pt)", "Median (es)"),
)
_LANGUAGES = ("English", "French", "Portuguese", "Spanish")


def _bindings(
    *,
    languages: list[int],
    matrix: bool,
    english_range: str = "phrases!C2:C3",
) -> str:
    enum = ", ".join(str(item) for item in languages)
    language = f"""\
- id: language_index
  sheet: lang
  data_range: lang!A1
  layout: scalar
  input:
    domain:
      enum: [{enum}]
  structure:
    measure:
      concept: OBS_VALUE
      dtype: int
      bind:
        kind: data_cell
        read: int
    dimensions: []
  key: []
"""
    if matrix:
        phrases = """\
- id: phrase_by_language
  sheet: phrases
  data_range: phrases!C2:F3
  layout: matrix
  constant: {}
  structure:
    measure:
      concept: OBS_VALUE
      dtype: string
      bind:
        kind: data_cell
        read: string
    dimensions:
    - id: INDICATOR
      concept: INDICATOR
      role: key
      scope: cell
      bind:
        kind: row_label
        label_column: C
        read: string
        normalize: strip
    - id: LANGUAGE
      concept: LANGUAGE
      role: key
      scope: cell
      bind:
        kind: column_header
        header_row: 1
        read: string
        normalize: strip
  key:
  - INDICATOR
  - LANGUAGE
"""
    else:
        phrases = """\
- id: english_phrase
  sheet: phrases
  data_range: {english_range}
  layout: series
  constant: {}
  structure:
    measure:
      concept: OBS_VALUE
      dtype: string
      bind:
        kind: data_cell
        read: string
    dimensions:
    - id: INDICATOR
      concept: INDICATOR
      role: key
      scope: cell
      bind:
        kind: data_cell
        read: string
  key:
  - INDICATOR
- id: translated_phrase
  sheet: phrases
  data_range: phrases!D2:F3
  layout: matrix
  constant: {}
  structure:
    measure:
      concept: OBS_VALUE
      dtype: string
      bind:
        kind: data_cell
        read: string
    dimensions:
    - id: INDICATOR
      concept: INDICATOR
      role: key
      scope: cell
      bind:
        kind: row_label
        label_column: C
        read: string
        normalize: strip
    - id: LANGUAGE
      concept: LANGUAGE
      role: key
      scope: cell
      bind:
        kind: column_header
        header_row: 1
        read: string
        normalize: strip
  key:
  - INDICATOR
  - LANGUAGE
""".replace("{english_range}", english_range)
    return f"""\
schema_version: 1.21.0
workbook: mcve.xlsx
concept_scheme:
  id: mcve
  concepts:
  - id: OBS_VALUE
    name: Observation value
    dtype: string
  - id: INDICATOR
    name: Indicator
    dtype: string
  - id: LANGUAGE
    name: Language
    dtype: string
series:
{language}{phrases}- id: selected_phrase
  sheet: out
  data_range: out!B1
  layout: scalar
  output:
    compute:
      name: compute_selected_phrase
      record_contract: records
  structure:
    measure:
      concept: OBS_VALUE
      dtype: string
      bind:
        kind: data_cell
        read: string
    dimensions: []
  key: []
"""


def _write_workbook(path: Path, formula: str) -> None:
    from fastpyxl import Workbook

    workbook = Workbook()
    phrases = workbook.active
    assert phrases is not None
    phrases.title = "phrases"
    for index, language in enumerate(_LANGUAGES):
        phrases.cell(1, 3 + index, language)
    for row, values in enumerate(_PHRASES, start=2):
        for index, value in enumerate(values):
            phrases.cell(row, 3 + index, value)
    lang = workbook.create_sheet("lang")
    lang["A1"] = 0
    out = workbook.create_sheet("out")
    out["B1"] = formula
    workbook.save(path)


def _export(
    directory: Path,
    *,
    formula: str,
    languages: list[int],
    matrix: bool,
    provenance: bool,
    english_range: str = "phrases!C2:C3",
) -> tuple[dict[str, str], object]:
    workbook_path = directory / "mcve.xlsx"
    bindings_path = directory / "bindings.yaml"
    _write_workbook(workbook_path, formula)
    bindings_path.write_text(
        _bindings(languages=languages, matrix=matrix, english_range=english_range),
        encoding="utf-8",
    )
    bindings = load_series_bindings(bindings_path)
    graph = create_dependency_graph(
        workbook_path,
        ["out!B1"],
        dynamic_refs=DynamicRefConfig.from_bindings(
            bindings,
            workbook_path,
            bindings_path=bindings_path,
        ),
        capture_dependency_provenance=provenance,
    )
    modules = CodeGenerator(graph).generate_modules(
        series_bindings=bindings,
        bindings_workbook=workbook_path,
    )
    return modules, graph


def _selected(pkg: object, language_index: int) -> object:
    return invoke_public_compute(
        pkg,
        pkg.compute_selected_phrase,  # type: ignore[attr-defined]
        {"language_index": language_index},
    )


@pytest.mark.parametrize("provenance", [False, True])
def test_column_offset_steps_from_anchor_key(tmp_path: Path, provenance: bool) -> None:
    modules, graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        languages=[0, 1, 2, 3],
        matrix=False,
        provenance=provenance,
    )
    source = modules["internals.py"]
    assert "axis_step" in source
    assert "'English'" in source
    assert "english_phrase" in source
    assert "translated_phrase" in source
    pkg = load_package(modules, tmp_path, name=f"col_off_{int(provenance)}")
    evaluator = FormulaEvaluator(graph)
    for index in range(len(_LANGUAGES)):
        assert _selected(pkg, index) == _PHRASES[1][index]
        graph.set_node_value("lang!A1", index)
        evaluated = evaluator.evaluate(["out!B1"])["out!B1"]
        assert evaluated == _PHRASES[1][index]


def test_gapped_language_domain_keeps_worksheet_columns(tmp_path: Path) -> None:
    modules, _graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        languages=[0, 2],
        matrix=False,
        provenance=False,
    )
    source = modules["internals.py"]
    assert "('English', 'French', 'Portuguese')" in source
    pkg = load_package(modules, tmp_path, name="col_gap")
    assert _selected(pkg, 0) == "Median"
    assert _selected(pkg, 2) == "Median (pt)"


def test_literal_column_offset_reads_landing_series(tmp_path: Path) -> None:
    modules, _graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,2)",
        languages=[0, 1, 2, 3],
        matrix=False,
        provenance=False,
    )
    source = modules["internals.py"]
    assert "translated_phrase" in source
    assert "Portuguese" in source
    pkg = load_package(modules, tmp_path, name="col_lit")
    assert invoke_public_compute(pkg, pkg.compute_selected_phrase, {}) == "Median (pt)"


def test_matrix_binding_still_steps_its_column_axis(tmp_path: Path) -> None:
    modules, _graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        languages=[0, 1, 2, 3],
        matrix=True,
        provenance=False,
    )
    source = modules["internals.py"]
    assert (
        "phrase_by_language['Median', axis_step(data.LANGUAGE_AXIS, 'English', language_index)]"
        in source
    )
    pkg = load_package(modules, tmp_path, name="col_matrix")
    for index in range(4):
        assert _selected(pkg, index) == _PHRASES[1][index]


def test_one_cell_anchor_steps_from_anchor_key(tmp_path: Path) -> None:
    modules, graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        languages=[0, 1, 2, 3],
        matrix=False,
        provenance=False,
        english_range="phrases!C3",
    )
    source = modules["internals.py"]
    assert "axis_step" in source
    assert "'English'" in source
    assert "english_phrase['Median']" in source
    assert "translated_phrase" in source
    pkg = load_package(modules, tmp_path, name="col_one")
    evaluator = FormulaEvaluator(graph)
    for index in range(len(_LANGUAGES)):
        assert _selected(pkg, index) == _PHRASES[1][index]
        graph.set_node_value("lang!A1", index)
        evaluated = evaluator.evaluate(["out!B1"])["out!B1"]
        assert evaluated == _PHRASES[1][index]


def test_one_cell_literal_column_offset_reads_landing_series(tmp_path: Path) -> None:
    modules, _graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,2)",
        languages=[0, 1, 2, 3],
        matrix=False,
        provenance=False,
        english_range="phrases!C3",
    )
    source = modules["internals.py"]
    assert "translated_phrase" in source
    assert "Portuguese" in source
    assert "axis_step" not in source
    pkg = load_package(modules, tmp_path, name="col_one_lit")
    assert invoke_public_compute(pkg, pkg.compute_selected_phrase, {}) == "Median (pt)"


def test_unbound_column_offset_fails_closed(tmp_path: Path) -> None:
    text = _bindings(languages=[0, 1, 2, 3], matrix=False)
    document = yaml.safe_load(text)
    document["series"] = [
        series for series in document["series"] if series["id"] != "translated_phrase"
    ]
    workbook_path = tmp_path / "mcve.xlsx"
    bindings_path = tmp_path / "bindings.yaml"
    _write_workbook(workbook_path, "=OFFSET(phrases!C3,0,lang!A1)")
    bindings_path.write_text(yaml.safe_dump(document), encoding="utf-8")
    bindings = load_series_bindings(bindings_path)
    graph = create_dependency_graph(
        workbook_path,
        ["out!B1"],
        dynamic_refs=DynamicRefConfig.from_bindings(
            bindings,
            workbook_path,
            bindings_path=bindings_path,
        ),
    )
    with pytest.raises(InvertedTreeExportError, match="column offset"):
        CodeGenerator(graph).generate_modules(
            series_bindings=bindings,
            bindings_workbook=workbook_path,
        )
