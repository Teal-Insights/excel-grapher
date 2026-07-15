"""Contract tests for INDEX/OFFSET reference pairing in export runtime."""

from __future__ import annotations

import typing

from excel_grapher.exporter.export_runtime.offset import (
    OffsetRefInfo,
    xl_index_ref,
    xl_offset,
    xl_offset_ref,
)
from excel_grapher.exporter.export_runtime.values import ExcelRange
from excel_grapher.runtime.cache import EvalContext


def test_xl_index_ref_returns_address_metadata_not_cell_value() -> None:
    values = {
        "Sheet1!B2": 42.0,
        "Sheet1!B3": 99.0,
        "Sheet1!C2": 7.0,
    }

    def resolver(address: str):
        if address in values:
            return lambda ctx: values[address]
        return None

    ctx = EvalContext(inputs={}, resolver=resolver)

    ref = xl_index_ref(("Sheet1", 2, 2, 3, 3), 1.0, 1.0)
    assert ref == ("Sheet1", 2, 2)
    assert xl_offset(ctx, ref, 0.0, 0.0) == 42.0


def test_xl_index_ref_and_xl_offset_docstrings_describe_pairing() -> None:
    assert xl_index_ref.__doc__ is not None
    assert "not a cell value" in xl_index_ref.__doc__.lower()
    assert "xl_offset" in xl_index_ref.__doc__

    assert xl_offset.__doc__ is not None
    assert "xl_index_ref" in xl_offset.__doc__
    assert "cell value" in xl_offset.__doc__.lower()


def test_offset_ref_info_annotations_use_named_alias_not_cell_value() -> None:
    index_hints = typing.get_type_hints(xl_index_ref)
    offset_hints = typing.get_type_hints(xl_offset)
    offset_ref_hints = typing.get_type_hints(xl_offset_ref)

    assert index_hints["return"] is OffsetRefInfo
    assert offset_hints["ref_info"] is OffsetRefInfo
    assert offset_hints["return"] is not OffsetRefInfo
    assert xl_offset.__annotations__["return"] == "CellValue"
    assert offset_ref_hints["return"] is ExcelRange

    # Raise-only export API: refs are address metadata, not error sentinels or values.
    assert "XlError" not in xl_index_ref.__annotations__["return"]
    assert "XlError" not in xl_offset.__annotations__["ref_info"]
    assert "CellValue" not in xl_index_ref.__annotations__["return"]
