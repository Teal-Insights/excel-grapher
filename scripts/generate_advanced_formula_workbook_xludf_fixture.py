"""Generate ``examples/micro_workbooks/advanced_formula_workbook_xludf`` fixtures."""

from __future__ import annotations

from pathlib import Path

from tests.integration.utils.rewrite_xludf_workbook import write_xludf_workbook_copy

REPO = Path(__file__).resolve().parents[1]
SANDBOX = REPO / "sandbox" / "model"
EXAMPLES = REPO / "examples" / "micro_workbooks"
SOURCE_WB = SANDBOX / "advanced_formula_workbook.xlsx"
SOURCE_BINDINGS = SANDBOX / "advanced_formula_workbook.bindings"
DEST_WB = EXAMPLES / "advanced_formula_workbook_xludf.xlsx"
DEST_BINDINGS = EXAMPLES / "advanced_formula_workbook_xludf.bindings"


def main() -> None:
    """Copy sandbox workbook/bindings and rewrite formulas to ``_xludf.`` spelling."""
    if not SOURCE_WB.is_file():
        raise SystemExit(f"Missing source workbook: {SOURCE_WB}")
    if not SOURCE_BINDINGS.is_dir():
        raise SystemExit(f"Missing source bindings: {SOURCE_BINDINGS}")

    DEST_BINDINGS.mkdir(parents=True, exist_ok=True)
    for shard in SOURCE_BINDINGS.glob("*.bindings.yaml"):
        text = shard.read_text(encoding="utf-8").replace(
            "workbook: advanced_formula_workbook.xlsx",
            "workbook: advanced_formula_workbook_xludf.xlsx",
        )
        (DEST_BINDINGS / shard.name).write_text(text, encoding="utf-8")

    write_xludf_workbook_copy(SOURCE_WB, DEST_WB)
    print(f"wrote {DEST_WB}")
    print(f"wrote {len(list(DEST_BINDINGS.glob('*.bindings.yaml')))} binding shard(s)")


if __name__ == "__main__":
    main()
