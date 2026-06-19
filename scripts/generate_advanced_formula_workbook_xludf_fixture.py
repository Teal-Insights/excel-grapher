"""Generate sandbox ``advanced_formula_workbook`` prefix-variant fixtures.

Writes ``_xludf`` and ``_xlfn`` workbook copies plus matching binding shards under
``sandbox/model/``. These fixtures are gitignored and not part of the main repo.
"""

from __future__ import annotations

from pathlib import Path

from tests.integration.utils.rewrite_xludf_workbook import write_xludf_workbook_copy

REPO = Path(__file__).resolve().parents[1]
SANDBOX = REPO / "sandbox" / "model"
SOURCE_WB = SANDBOX / "advanced_formula_workbook.xlsx"
SOURCE_BINDINGS = SANDBOX / "advanced_formula_workbook.bindings"

VARIANTS: tuple[tuple[str, str, str], ...] = (
    ("xludf", "advanced_formula_workbook_xludf.xlsx", "advanced_formula_workbook_xludf.bindings"),
    ("xlfn", "advanced_formula_workbook_xlfn.xlsx", "advanced_formula_workbook_xlfn.bindings"),
)


def _write_binding_shards(dest_bindings: Path, *, workbook_name: str) -> int:
    dest_bindings.mkdir(parents=True, exist_ok=True)
    count = 0
    for shard in SOURCE_BINDINGS.glob("*.bindings.yaml"):
        text = shard.read_text(encoding="utf-8").replace(
            "workbook: advanced_formula_workbook.xlsx",
            f"workbook: {workbook_name}",
        )
        (dest_bindings / shard.name).write_text(text, encoding="utf-8")
        count += 1
    return count


def main() -> None:
    """Copy sandbox workbook/bindings and emit prefix-variant fixtures."""
    if not SOURCE_WB.is_file():
        raise SystemExit(f"Missing source workbook: {SOURCE_WB}")
    if not SOURCE_BINDINGS.is_dir():
        raise SystemExit(f"Missing source bindings: {SOURCE_BINDINGS}")

    for variant, workbook_name, bindings_dir_name in VARIANTS:
        dest_wb = SANDBOX / workbook_name
        dest_bindings = SANDBOX / bindings_dir_name
        shard_count = _write_binding_shards(dest_bindings, workbook_name=workbook_name)
        if variant == "xludf":
            write_xludf_workbook_copy(SOURCE_WB, dest_wb)
        else:
            dest_wb.write_bytes(SOURCE_WB.read_bytes())
        print(f"[{variant}] wrote {dest_wb}")
        print(f"[{variant}] wrote {shard_count} binding shard(s) under {dest_bindings}")


if __name__ == "__main__":
    main()
