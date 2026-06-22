"""Tiny DSA workbook fixture for similarity-aware compression (issue #282).

The workbook models six compressible subgraphs on the Engine sheet:

- **Groups 1–5** (``parallel_family="shocked_year_block"``): columns C–G at row 20.
  Each root ``Engine!{col}20`` shares the snowball ratio template
  ``B20*(1+{col}15/100)/(1+{col}14/100)-{col}16`` with column-specific CHOOSE
  intermediates at rows 14–16. Only the column letter and matching Inputs
  growth cells differ.
- **Group 6** (``parallel_family="linear_aggregate"``): ``Engine!H20`` uses a
  linear ``H14+H15-H16`` shape without CHOOSE or the shocked ratio template.

Shared baseline ``Engine!B20`` feeds all five shocked-year roots. Optimal
compression removes ``B20`` as an identity transit; similarity-aware candidates
still treat each column's rows 14–16 plus ``{col}20`` as one compressible group.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import fastpyxl

ParallelFamilyId = Literal["shocked_year_block", "linear_aggregate"]

SHOCKED_YEAR_COLUMNS: tuple[str, ...] = ("C", "D", "E", "F", "G")
GROUP_6_ROOT = "Engine!H20"
TINY_DSA_TARGETS: tuple[str, ...] = tuple(f"Engine!{col}20" for col in SHOCKED_YEAR_COLUMNS) + (
    GROUP_6_ROOT,
)


@dataclass(frozen=True)
class TinyDsaGroup:
    """One compressible subgraph in the Tiny DSA fixture."""

    group_id: int
    root: str
    members: frozenset[str]
    parallel_family: ParallelFamilyId

    @property
    def internal_members(self) -> frozenset[str]:
        """Nodes removed when collapsing into ``root`` (members minus root)."""
        return frozenset(m for m in self.members if m != self.root)


def shocked_year_group(column: str, *, group_id: int) -> TinyDsaGroup:
    """Build expected metadata for one shocked-year column block."""
    col = column.upper()
    root = f"Engine!{col}20"
    members = frozenset(
        {
            root,
            f"Engine!{col}14",
            f"Engine!{col}15",
            f"Engine!{col}16",
        }
    )
    return TinyDsaGroup(
        group_id=group_id,
        root=root,
        members=members,
        parallel_family="shocked_year_block",
    )


def group_6_linear() -> TinyDsaGroup:
    """Build expected metadata for the non-parallel linear aggregate block."""
    return TinyDsaGroup(
        group_id=6,
        root=GROUP_6_ROOT,
        members=frozenset(
            {
                GROUP_6_ROOT,
                "Engine!H14",
                "Engine!H15",
                "Engine!H16",
            }
        ),
        parallel_family="linear_aggregate",
    )


TINY_DSA_GROUPS: tuple[TinyDsaGroup, ...] = tuple(
    shocked_year_group(col, group_id=index + 1) for index, col in enumerate(SHOCKED_YEAR_COLUMNS)
) + (group_6_linear(),)

SHOCKED_YEAR_FAMILY_GROUPS: tuple[TinyDsaGroup, ...] = tuple(
    g for g in TINY_DSA_GROUPS if g.parallel_family == "shocked_year_block"
)
LINEAR_FAMILY_GROUPS: tuple[TinyDsaGroup, ...] = tuple(
    g for g in TINY_DSA_GROUPS if g.parallel_family == "linear_aggregate"
)


def _shocked_year_cells(
    column: str, *, year_index: int
) -> dict[tuple[str, str], str | int | float]:
    """Return Engine/Inputs cell entries for one shocked-year column."""
    col = column.upper()
    return {
        ("Inputs", f"{col}16"): 2.0 + year_index - 1,
        ("Inputs", f"{col}17"): 3.0 + year_index - 1,
        ("Inputs", f"{col}18"): -1.0,
        ("Engine", f"{col}5"): year_index,
        ("Engine", f"{col}10"): f"=IF({col}5>=Inputs!$B$21,1,0)",
        ("Engine", f"{col}14"): (f"=Inputs!{col}16+CHOOSE(Inputs!$B$22,$B$9,0,0)*{col}10"),
        ("Engine", f"{col}15"): (f"=Inputs!{col}17+CHOOSE(Inputs!$B$22,0,$B$9,0)*{col}10"),
        ("Engine", f"{col}16"): (f"=Inputs!{col}18+CHOOSE(Inputs!$B$22,0,0,$B$9)*{col}10"),
        ("Engine", f"{col}20"): f"=$B$20*(1+{col}15/100)/(1+{col}14/100)-{col}16",
    }


def build_tiny_dsa_workbook(path: Path) -> None:
    """Write the Tiny DSA xlsx fixture to ``path``."""
    wb = fastpyxl.Workbook()
    ws_inputs = wb.active
    ws_inputs.title = "Inputs"
    wb.create_sheet("Engine")
    ws_engine = wb["Engine"]

    cells: dict[tuple[str, str], str | int | float] = {
        ("Inputs", "B6"): 100,
        ("Inputs", "B21"): 1,
        ("Inputs", "B22"): 1,
        ("Engine", "B9"): 0.5,
        ("Engine", "B20"): "=Inputs!B6",
        ("Inputs", "H16"): 4.0,
        ("Inputs", "H17"): 1.5,
        ("Inputs", "H18"): 0.25,
        ("Engine", "H14"): "=Inputs!H16",
        ("Engine", "H15"): "=Inputs!H17",
        ("Engine", "H16"): "=Inputs!H18",
        ("Engine", "H20"): "=H14+H15-H16",
    }
    for index, col in enumerate(SHOCKED_YEAR_COLUMNS, start=1):
        cells.update(_shocked_year_cells(col, year_index=index))

    for (sheet, addr), value in cells.items():
        worksheet = ws_inputs if sheet == "Inputs" else ws_engine
        worksheet[addr] = value

    wb.save(path)
