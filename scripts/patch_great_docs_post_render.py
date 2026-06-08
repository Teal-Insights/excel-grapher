"""Patch great-docs post-render script for UTF-8 stdout on Windows.

great-docs copies `assets/post-render.py` into `great-docs/scripts/` on every
build. That script prints emoji status lines; Windows consoles often default to
a legacy code page, which raises `UnicodeEncodeError` during post-render.

This pre-render hook injects a small `sys.stdout` / `sys.stderr` reconfigure
block into `scripts/post-render.py` before Quarto runs it.
"""

from __future__ import annotations

from pathlib import Path

_MARKER = "_configure_gd_stdio"
_POST_RENDER = Path("scripts/post-render.py")
_INSERT_AFTER = "import re\n"
_INSERT_BLOCK = """import sys


def _configure_gd_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            pass


_configure_gd_stdio()
"""


def main() -> None:
    if not _POST_RENDER.is_file():
        raise SystemExit(f"post-render script not found: {_POST_RENDER}")

    text = _POST_RENDER.read_text(encoding="utf-8")
    if _MARKER in text:
        return

    if _INSERT_AFTER not in text:
        raise SystemExit(
            "great-docs post-render.py layout changed; update patch_great_docs_post_render.py"
        )

    _POST_RENDER.write_text(
        text.replace(_INSERT_AFTER, _INSERT_AFTER + _INSERT_BLOCK, 1), encoding="utf-8"
    )


if __name__ == "__main__":
    main()
