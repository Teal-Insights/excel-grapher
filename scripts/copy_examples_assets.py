"""Copy non-Markdown assets from examples/ into the great-docs build tree.

Great Docs custom sections copy only ``.qmd`` and ``.md`` files. Example
workbooks (``.xlsx``), images, and other co-located assets must be synced
before Quarto executes code cells that reference them.
"""

from __future__ import annotations

import shutil
from pathlib import Path

_SKIP_SUFFIXES = {".qmd", ".md"}


def _find_project_root(start: Path) -> Path:
    for candidate in (start, *start.parents):
        if (candidate / "pyproject.toml").is_file():
            return candidate
    return start


def main() -> None:
    build_dir = Path.cwd()
    project_root = _find_project_root(build_dir)
    src_root = project_root / "examples"
    dst_root = build_dir / "examples"

    if not src_root.is_dir() or not dst_root.is_dir():
        return

    copied = 0
    for path in src_root.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() in _SKIP_SUFFIXES:
            continue
        rel = path.relative_to(src_root)
        dest = dst_root / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, dest)
        copied += 1

    if copied:
        print(f"Copied {copied} example asset file(s) into {dst_root.relative_to(build_dir)}/")

    schema_src = project_root / "excel_grapher" / "series_bindings" / "series_binding.schema.json"
    if schema_src.is_file():
        shutil.copy2(schema_src, build_dir / "series_binding.schema.json")


if __name__ == "__main__":
    main()
