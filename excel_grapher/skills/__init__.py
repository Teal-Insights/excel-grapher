"""Locate the separately installed author-bindings skill."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from importlib.resources import as_file, files
from pathlib import Path

_INSTALL_HINT = (
    "The author-bindings skill is not part of the excel-grapher wheel. "
    "Install the excel-grapher-author-bindings distribution "
    "(uv add excel-grapher --extra skills) or copy skills/author-bindings "
    "from the source tree / sdist into .cursor/skills/author-bindings "
    "or Claude's skill path."
)


class AuthorBindingsSkillNotInstalledError(FileNotFoundError):
    """Raised when the author-bindings skill distribution is not available."""


def _source_tree_skill_dir() -> Path | None:
    repo_skill = Path(__file__).resolve().parents[2] / "skills" / "author-bindings"
    if (repo_skill / "SKILL.md").is_file():
        return repo_skill
    return None


@contextmanager
def author_bindings_skill_dir() -> Iterator[Path]:
    """Yield the author-bindings skill directory.

    The skill is a separate distribution artifact, not part of the
    `excel-grapher` wheel. Resolution order:

    1. Installed `excel_grapher_author_bindings` package data (`as_file`)
    2. Source-tree / sdist `skills/author-bindings`

    Yields:
        Filesystem path to the skill tree (contains `SKILL.md`).

    Raises:
        AuthorBindingsSkillNotInstalledError: Neither the extra package nor a
            source checkout is available.
    """
    try:
        traversable = files("excel_grapher_author_bindings").joinpath("author-bindings")
    except ModuleNotFoundError:
        traversable = None
    if traversable is not None:
        with as_file(traversable) as path:
            yielded = Path(path)
            if (yielded / "SKILL.md").is_file():
                yield yielded
                return
    source = _source_tree_skill_dir()
    if source is not None:
        yield source
        return
    raise AuthorBindingsSkillNotInstalledError(_INSTALL_HINT)
