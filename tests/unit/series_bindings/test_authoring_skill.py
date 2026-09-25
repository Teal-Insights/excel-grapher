"""Tests for the author-bindings agent skill artifact."""

from __future__ import annotations

import shutil
from pathlib import Path

import yaml

from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from excel_grapher.skills import author_bindings_skill_dir

_REPO_ROOT = Path(__file__).resolve().parents[3]
_CANONICAL = _REPO_ROOT / "skills" / "author-bindings"
_AGENTS_COPY = _REPO_ROOT / ".agents" / "skills" / "author-bindings"
_CURSOR_COPY = _REPO_ROOT / ".cursor" / "skills" / "author-bindings"
_EXTRA_PACKAGE_COPY = (
    _REPO_ROOT
    / "packages"
    / "excel-grapher-author-bindings"
    / "excel_grapher_author_bindings"
    / "author-bindings"
)


def _skill_files(root: Path) -> dict[str, str]:
    files: dict[str, str] = {}
    for path in sorted(root.rglob("*")):
        if path.is_file():
            files[str(path.relative_to(root))] = path.read_text(encoding="utf-8")
    return files


def _sync_skill_tree(source: Path, dest: Path) -> None:
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(source, dest)


def test_generate_agent_facing_skill_copies() -> None:
    """Sync the consumer package from the canonical tree.

    This repository does not install the skill for its own agents.
    """
    _sync_skill_tree(_CANONICAL, _EXTRA_PACKAGE_COPY)
    assert _skill_files(_EXTRA_PACKAGE_COPY) == _skill_files(_CANONICAL)
    assert not _AGENTS_COPY.exists()
    assert not _CURSOR_COPY.exists()


def test_skill_install_is_copy_into_agents_skills() -> None:
    from excel_grapher.skills import _INSTALL_HINT

    skill = (_CANONICAL / "SKILL.md").read_text(encoding="utf-8")
    guide = (_REPO_ROOT / "user_guide" / "09-binding-authoring.qmd").read_text(encoding="utf-8")
    extra_readme = (
        _REPO_ROOT / "packages" / "excel-grapher-author-bindings" / "README.md"
    ).read_text(encoding="utf-8")
    root_readme = (_REPO_ROOT / "README.md").read_text(encoding="utf-8")
    for text in (skill, guide, extra_readme, root_readme, _INSTALL_HINT):
        assert ".agents/skills/author-bindings" in text
        assert ".cursor/skills/author-bindings" not in text
        assert "uv add excel-grapher --extra skills" not in text
        assert "excel-grapher[skills]" not in text
    assert "mkdir -p .agents/skills" in skill
    assert "cp -R skills/author-bindings .agents/skills/author-bindings" in skill
    assert "cp -R skills/author-bindings .agents/skills/author-bindings" in guide
    assert "cp -R skills/author-bindings .agents/skills/author-bindings" in root_readme


def test_library_package_does_not_ship_skill_markdown() -> None:
    import excel_grapher.skills as skills_pkg

    packaged = Path(skills_pkg.__file__).resolve().parent / "author-bindings"
    assert not (packaged / "SKILL.md").is_file()


def test_author_bindings_skill_dir_uses_as_file_context() -> None:
    with author_bindings_skill_dir() as packaged:
        assert packaged.is_dir()
        assert (packaged / "SKILL.md").is_file()
        assert _skill_files(packaged) == _skill_files(_CANONICAL)


def test_skill_documents_authoring_loop_without_tiny_dsa() -> None:
    texts = [
        (_CANONICAL / "SKILL.md").read_text(encoding="utf-8"),
        (_CANONICAL / "references" / "conventions.md").read_text(encoding="utf-8"),
        (_CANONICAL / "references" / "pitfalls.md").read_text(encoding="utf-8"),
        (_CANONICAL / "references" / "validation-loop.md").read_text(encoding="utf-8"),
    ]
    joined = "\n".join(texts)
    lowered = joined.lower()
    assert "CURRENT_SCHEMA_VERSION" in joined
    assert "fill: true" in joined
    assert "constant: {}" in joined
    assert "bind.kind: constant" in joined
    assert "unreachable" in lowered
    assert "uniquify" in lowered or "unique" in lowered
    assert "matrix" in lowered
    assert "semantic family" in lowered
    assert "coverage worklist" in lowered
    assert "bound closure" in lowered or "bound graph closure" in lowered
    assert "full workbook walk" in lowered
    assert "upsert" in lowered
    assert "tiny dsa" not in lowered
    assert "tiny-dsa" not in lowered
    assert "workbook_config" not in lowered
    assert "load_pipeline_config" not in lowered
    assert "bindings candidates" in joined
    assert "undomained" in lowered
    assert "use_cached_dynamic_refs=False" in joined
    assert "SeriesRelationError" in joined
    assert "BLANK_RANGES" in joined
    assert "partial_graph_overlap" in joined
    assert "leaf_in_formula_series" in joined
    assert "non_leaf_input_overlap" in joined
    assert "non_leaf_constant_overlap" in joined
    assert "no_formula_override_targets" in joined
    assert "no_leaf_constant_targets" in joined
    assert "no_formula_internal_targets" in joined
    assert "data_range has no graph formula cells" in joined
    assert "intersect_graph_leaves: false" in joined
    assert "axis_labels" in joined
    assert "do not narrow" in lowered
    assert "must be on-graph" not in lowered
    assert "bindings emit" not in lowered
    assert "author_bindings" not in lowered


def test_measure_shard_example_keeps_fill_and_shared_compute_name() -> None:
    example = yaml.safe_load(
        (_CANONICAL / "assets" / "measure-shards.example.yaml").read_text(encoding="utf-8")
    )
    assert example["schema_version"] == CURRENT_SCHEMA_VERSION
    outputs = example["outputs"]["series"]
    assert {series["output"]["compute"]["name"] for series in outputs} == {"compute_gap_milestones"}
    for series in outputs:
        time_binds = [
            dim["bind"] for dim in series["structure"]["dimensions"] if dim["id"] == "TIME_PERIOD"
        ]
        assert time_binds[0]["fill"] is True
    internals = example["internals"]["series"]
    assert set(internals[0]["key"]) == {"SCENARIO", "TIME_PERIOD", "MEASURE"}


def test_catalog_example_is_pedagogical_not_an_emitter_input() -> None:
    example = yaml.safe_load(
        (_CANONICAL / "assets" / "catalog.example.yaml").read_text(encoding="utf-8")
    )
    assert example["schema_version"] == CURRENT_SCHEMA_VERSION
    for direction in ("inputs", "outputs", "internals", "constants"):
        assert isinstance(example[direction]["series"], list)
    skill = (_CANONICAL / "SKILL.md").read_text(encoding="utf-8")
    assert "not an input to a generator" in skill.lower() or "pedagogical" in skill.lower()


def test_sdist_includes_author_bindings_skill() -> None:
    pyproject = (_REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert "skills/author-bindings" in pyproject
    assert "excel-grapher-author-bindings" in pyproject
