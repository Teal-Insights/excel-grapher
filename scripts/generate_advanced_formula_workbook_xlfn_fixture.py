"""Generate ``examples/micro_workbooks/advanced_formula_workbook_xlfn`` fixtures."""

from __future__ import annotations

from pathlib import Path

from tests.integration.utils.rewrite_prefixed_workbook import write_xlfn_workbook_copy

REPO = Path(__file__).resolve().parents[1]
SANDBOX = REPO / "sandbox" / "model"
EXAMPLES = REPO / "examples" / "micro_workbooks"
SOURCE_WB = SANDBOX / "advanced_formula_workbook.xlsx"
SOURCE_BINDINGS = SANDBOX / "advanced_formula_workbook.bindings"
FALLBACK_WB = EXAMPLES / "advanced_formula_workbook_xludf.xlsx"
FALLBACK_BINDINGS = EXAMPLES / "advanced_formula_workbook_xludf.bindings"
DEST_WB = EXAMPLES / "advanced_formula_workbook_xlfn.xlsx"
DEST_BINDINGS = EXAMPLES / "advanced_formula_workbook_xlfn.bindings"


def main() -> None:
    """Copy workbook/bindings and rewrite formulas to ``_xlfn.`` spelling."""
    if SOURCE_WB.is_file() and SOURCE_BINDINGS.is_dir():
        source_wb = SOURCE_WB
        source_bindings = SOURCE_BINDINGS
        bindings_workbook_name = "advanced_formula_workbook.xlsx"
    elif FALLBACK_WB.is_file() and FALLBACK_BINDINGS.is_dir():
        source_wb = FALLBACK_WB
        source_bindings = FALLBACK_BINDINGS
        bindings_workbook_name = "advanced_formula_workbook_xludf.xlsx"
    else:
        raise SystemExit(
            "Missing source workbook/bindings. Expected sandbox advanced_formula_workbook "
            "or examples advanced_formula_workbook_xludf fixtures."
        )

    DEST_BINDINGS.mkdir(parents=True, exist_ok=True)
    for shard in source_bindings.glob("*.bindings.yaml"):
        text = shard.read_text(encoding="utf-8").replace(
            f"workbook: {bindings_workbook_name}",
            "workbook: advanced_formula_workbook_xlfn.xlsx",
        )
        (DEST_BINDINGS / shard.name).write_text(text, encoding="utf-8")

    write_xlfn_workbook_copy(source_wb, DEST_WB)
    print(f"wrote {DEST_WB}")
    print(f"wrote {len(list(DEST_BINDINGS.glob('*.bindings.yaml')))} binding shard(s)")


if __name__ == "__main__":
    main()
