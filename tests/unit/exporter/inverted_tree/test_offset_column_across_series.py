"""Column OFFSET from a series with no column axis (#991).

A one-column anchor and the translation block beside it are separate series.
`OFFSET(anchor, 0, language)` steps the column key that covers the in-domain
landings, starting at the anchor column.
"""

from __future__ import annotations

import types
from pathlib import Path
from typing import Any, cast

import pytest
import yaml

from excel_grapher import CodeGenerator, DynamicRefConfig, create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher.graph import DependencyGraph
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


def _write_workbook(
    path: Path,
    formula: str,
    *,
    headers: tuple[object, ...] | None = None,
) -> None:
    from fastpyxl import Workbook

    workbook = Workbook()
    phrases = workbook.active
    assert phrases is not None
    phrases.title = "phrases"
    for index, language in enumerate(headers if headers is not None else _LANGUAGES):
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
) -> tuple[dict[str, str], DependencyGraph]:
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


def _selected(pkg: types.ModuleType, language_index: int) -> object:
    return invoke_public_compute(
        pkg,
        pkg.compute_selected_phrase,
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
    assert (
        "axis_step(('English', 'French', 'Portuguese', 'Spanish'), 'English', language_index)"
        in source
    )
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


@pytest.mark.parametrize("provenance", [False, True])
def test_one_cell_anchor_steps_from_anchor_key(tmp_path: Path, provenance: bool) -> None:
    modules, graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        languages=[0, 1, 2, 3],
        matrix=False,
        provenance=provenance,
        english_range="phrases!C3",
    )
    source = modules["internals.py"]
    assert "axis_step" in source
    assert "'English'" in source
    assert "english_phrase['Median']" in source
    assert "translated_phrase" in source
    pkg = load_package(modules, tmp_path, name=f"col_one_{int(provenance)}")
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
    with pytest.raises(InvertedTreeExportError, match="unbound"):
        CodeGenerator(graph).generate_modules(
            series_bindings=bindings,
            bindings_workbook=workbook_path,
        )


def _export_document(
    directory: Path,
    document: dict[str, Any],
    *,
    formula: str,
    provenance: bool = False,
    headers: tuple[object, ...] | None = None,
) -> tuple[dict[str, str], DependencyGraph]:
    workbook_path = directory / "mcve.xlsx"
    bindings_path = directory / "bindings.yaml"
    _write_workbook(workbook_path, formula, headers=headers)
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
        capture_dependency_provenance=provenance,
    )
    modules = CodeGenerator(graph).generate_modules(
        series_bindings=bindings,
        bindings_workbook=workbook_path,
    )
    return modules, graph


def _split_document(
    *, languages: list[int], english_range: str = "phrases!C2:C3"
) -> dict[str, Any]:
    loaded = yaml.safe_load(
        _bindings(languages=languages, matrix=False, english_range=english_range)
    )
    if not isinstance(loaded, dict):
        raise AssertionError("bindings document is not a mapping")
    return cast(dict[str, Any], loaded)


@pytest.mark.parametrize("languages", [[2], [1, 2, 3]])
def test_domain_without_zero_keeps_the_anchor_column(tmp_path: Path, languages: list[int]) -> None:
    modules, _graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        languages=languages,
        matrix=False,
        provenance=False,
    )
    source = modules["internals.py"]
    if languages == [2]:
        expected = "axis_step(('English', 'French', 'Portuguese'), 'English', language_index)"
    else:
        expected = (
            "axis_step(('English', 'French', 'Portuguese', 'Spanish'), 'English', language_index)"
        )
    assert expected in source
    pkg = load_package(modules, tmp_path, name=f"col_no_zero_{languages[0]}")
    if 2 in languages:
        assert _selected(pkg, 2) == "Median (pt)"
    if 1 in languages:
        assert _selected(pkg, 1) == "Median (fr)"


def test_zero_only_domain_reads_the_anchor(tmp_path: Path) -> None:
    modules, _graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        languages=[0],
        matrix=False,
        provenance=False,
    )
    source = modules["internals.py"]
    assert "english_phrase['Median']" in source
    assert "axis_step" not in source
    pkg = load_package(modules, tmp_path, name="col_zero")
    assert _selected(pkg, 0) == "Median"


def test_interior_unbound_column_fails_closed(tmp_path: Path) -> None:
    document = _split_document(languages=[0, 2])
    for series in document["series"]:
        assert isinstance(series, dict)
        if series["id"] == "translated_phrase":
            series["data_range"] = "phrases!E2:F3"
    with pytest.raises(InvertedTreeExportError, match=r"unbound cell phrases!D3"):
        _export_document(
            tmp_path,
            document,
            formula="=OFFSET(phrases!C3,0,lang!A1)",
        )


def test_duplicate_column_key_fails_closed(tmp_path: Path) -> None:
    document = _split_document(languages=[0, 1, 2, 3])
    with pytest.raises(InvertedTreeExportError, match="shared"):
        _export_document(
            tmp_path,
            document,
            formula="=OFFSET(phrases!C3,0,lang!A1)",
            headers=("English", "English", "Portuguese", "Spanish"),
        )


def test_one_cell_literal_into_unbound_cell_fails_closed(tmp_path: Path) -> None:
    document = _split_document(languages=[0, 1, 2, 3], english_range="phrases!C3")
    with pytest.raises(InvertedTreeExportError, match=r"unbound cell phrases!B3"):
        _export_document(
            tmp_path,
            document,
            formula="=OFFSET(phrases!C3,0,-1)",
        )


def test_one_cell_without_a_landing_column_key_fails_closed(tmp_path: Path) -> None:
    document = _split_document(languages=[0, 1, 2, 3], english_range="phrases!C3")
    series = [
        entry
        for entry in document["series"]
        if isinstance(entry, dict) and entry["id"] != "translated_phrase"
    ]
    for column, series_id in (
        (4, "french_phrase"),
        (5, "portuguese_phrase"),
        (6, "spanish_phrase"),
    ):
        letter = "DEF"[column - 4]
        series.append(
            {
                "id": series_id,
                "sheet": "phrases",
                "data_range": f"phrases!{letter}3",
                "layout": "scalar",
                "constant": {},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "string",
                        "bind": {"kind": "data_cell", "read": "string"},
                    },
                    "dimensions": [],
                },
                "key": [],
            }
        )
    document["series"] = series
    with pytest.raises(InvertedTreeExportError, match="no column key"):
        _export_document(
            tmp_path,
            document,
            formula="=OFFSET(phrases!C3,0,lang!A1)",
        )


def test_one_cell_disagreeing_column_binds_fail_closed(tmp_path: Path) -> None:
    document = _split_document(languages=[0, 1, 2, 3], english_range="phrases!C3")
    translated = next(
        entry
        for entry in document["series"]
        if isinstance(entry, dict) and entry["id"] == "translated_phrase"
    )
    assert isinstance(translated, dict)
    french = yaml.safe_load(yaml.safe_dump(translated))
    rest = yaml.safe_load(yaml.safe_dump(translated))
    assert isinstance(french, dict)
    assert isinstance(rest, dict)
    french["id"] = "french_phrase"
    french["data_range"] = "phrases!D2:D3"
    rest["id"] = "later_phrase"
    rest["data_range"] = "phrases!E2:F3"
    for dimension in rest["structure"]["dimensions"]:
        if dimension["id"] == "LANGUAGE":
            dimension["bind"]["header_row"] = 2
    document["series"] = [
        entry
        for entry in document["series"]
        if isinstance(entry, dict) and entry["id"] != "translated_phrase"
    ]
    document["series"].extend([french, rest])
    with pytest.raises(InvertedTreeExportError, match="column key"):
        _export_document(
            tmp_path,
            document,
            formula="=OFFSET(phrases!C3,0,lang!A1)",
        )


def test_omitted_read_uses_the_landing_series_key_type(tmp_path: Path) -> None:
    document = _split_document(languages=[0, 1, 2, 3])
    for series in document["series"]:
        assert isinstance(series, dict)
        if series["id"] != "translated_phrase":
            continue
        for dimension in series["structure"]["dimensions"]:
            if dimension["id"] == "LANGUAGE":
                del dimension["bind"]["read"]
    modules, _graph = _export_document(
        tmp_path,
        document,
        formula="=OFFSET(phrases!C3,0,lang!A1)",
        headers=(0, 1, 2, 3),
    )
    source = modules["internals.py"]
    assert "axis_step(('0', '1', '2', '3'), '0', language_index)" in source
    pkg = load_package(modules, tmp_path, name="col_read")
    assert _selected(pkg, 0) == "Median"
    assert _selected(pkg, 2) == "Median (pt)"


def test_one_cell_row_move_stays_on_the_anchor(tmp_path: Path) -> None:
    modules, _graph = _export(
        tmp_path,
        formula="=OFFSET(phrases!C3,lang!A1,0)",
        languages=[0, 1],
        matrix=False,
        provenance=False,
        english_range="phrases!C3",
    )
    source = modules["internals.py"]
    assert "at_anchor" in source
    assert "axis_step" not in source
    pkg = load_package(modules, tmp_path, name="col_row")
    assert _selected(pkg, 1) == "#VALUE!"
