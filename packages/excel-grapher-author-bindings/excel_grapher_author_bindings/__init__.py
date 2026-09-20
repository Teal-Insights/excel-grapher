"""Separately installed author-bindings agent skill."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from importlib.resources import as_file, files
from pathlib import Path


@contextmanager
def author_bindings_skill_dir() -> Iterator[Path]:
    """Yield the packaged `author-bindings` skill directory.

    Uses `importlib.resources.as_file` so the tree is usable from a wheel or
    zipimport install.

    Yields:
        Filesystem path containing `SKILL.md`.
    """
    with as_file(files("excel_grapher_author_bindings").joinpath("author-bindings")) as path:
        yield Path(path)
