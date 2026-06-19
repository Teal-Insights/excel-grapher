"""Minimal workbooks and binding shards for ``_xlfn`` / ``_xludf`` prefix tests."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import xlsxwriter
import yaml

PrefixVariant = Literal["bare", "xlfn", "xludf"]

_LOOKUP_FORMULAS: dict[PrefixVariant, str] = {
    "bare": '=IFNA(Lookups!B1,"x")',
    "xlfn": '=_xlfn.IFNA(_xlfn.XLOOKUP(1,Lookups!A1:Lookups!A1,Lookups!B1:Lookups!B1),"x")',
    # Keep ``xludf`` free of XLOOKUP so Excel does not inject an extra ``_xlfn.`` prefix on save.
    "xludf": '=_xludf.IFNA(Lookups!B1,"x")',
}

_WORKBOOK_NAMES: dict[PrefixVariant, str] = {
    "bare": "prefix_bare.xlsx",
    "xlfn": "prefix_xlfn.xlsx",
    "xludf": "prefix_xludf.xlsx",
}


def workbook_filename(variant: PrefixVariant) -> str:
    """Return the workbook filename used by binding shards for ``variant``."""
    return _WORKBOOK_NAMES[variant]


def write_prefix_variant_workbook(path: Path, *, variant: PrefixVariant) -> Path:
    """Write a one-formula workbook using bare, ``_xlfn.``, or ``_xludf.`` spelling."""
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet("Lookups")
    worksheet.write_number(0, 0, 1)
    worksheet.write_string(0, 1, "hit")
    worksheet.write_formula(0, 2, _LOOKUP_FORMULAS[variant], None, "hit")
    workbook.close()
    return path


def prefix_variant_binding_document(*, variant: PrefixVariant) -> dict[str, Any]:
    """Minimal binding manifest referencing the variant workbook filename."""
    return {
        "schema_version": "1.3.0",
        "workbook": workbook_filename(variant),
        "series": [
            {
                "id": "lookup_result",
                "sheet": "Lookups",
                "data_range": "Lookups!C1",
                "layout": "scalar",
                "editable": False,
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "string",
                        "bind": {"kind": "data_cell", "read": "string"},
                    },
                    "dimensions": [],
                },
                "key": [],
                "output": {"compute": {"name": "compute_lookup_result"}},
            }
        ],
    }


def write_prefix_variant_binding_shard(
    path: Path,
    *,
    variant: PrefixVariant,
    series_id: str = "lookup_result",
) -> None:
    """Write one ``*.bindings.yaml`` shard for ``variant``."""
    document = prefix_variant_binding_document(variant=variant)
    document["series"][0]["id"] = series_id
    document["series"][0]["output"] = {
        "compute": {"name": f"compute_{series_id}"},
    }
    path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")
