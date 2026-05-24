"""Tests for input group contracts and record validation."""

from __future__ import annotations

import pytest

from excel_grapher.core.address_keys import normalize_key
from excel_grapher.exporter.input_groups import (
    BoundingBox,
    GroupingOverride,
    InputCell,
    InputGroup,
    InputGroupsPayload,
    InputGroupsSummary,
    SetterGenerationOptions,
    validate_record,
    validate_records,
)


class TestInputCellValidation:
    def test_canonical_address_required(self) -> None:
        with pytest.raises(ValueError, match="sheet-qualified"):
            InputCell(
                address="A1",
                sheet="Sheet1",
                row=1,
                col=1,
                row_labels=(),
                column_labels=(),
            )

    def test_components_must_match_address(self) -> None:
        with pytest.raises(ValueError, match="mismatch"):
            InputCell(
                address="Sheet1!A1",
                sheet="Sheet1",
                row=2,
                col=1,
                row_labels=(),
                column_labels=(),
            )

    def test_builds_from_address(self) -> None:
        cell = InputCell.from_address("Sheet1!B3", row_labels=("x",), column_labels=())
        assert cell.address == "Sheet1!B3"
        assert cell.sheet == "Sheet1"
        assert cell.row == 3
        assert cell.col == 2
        assert cell.row_labels == ("x",)


class TestRecordValidation:
    def test_requires_value(self) -> None:
        with pytest.raises(ValueError, match="value"):
            validate_record({})

    def test_accepts_minimal_record(self) -> None:
        validate_record({"value": 1.0})

    def test_validates_records_list(self) -> None:
        validate_records([{"value": 1}, {"value": "x"}])


class TestGroupingOverride:
    def test_range_must_be_sheet_qualified(self) -> None:
        with pytest.raises(ValueError, match="sheet-qualified"):
            GroupingOverride(range_spec="A1:B2", orientation="rowwise")


class TestInputGroupsPayloadRoundTrip:
    def test_to_dict_from_dict(self) -> None:
        cell = InputCell.from_address("S!A1")
        group = InputGroup(
            group_id="s_a1",
            sheet="S",
            orientation="rowwise",
            row_labels_key=(),
            column_labels_key=(),
            cells=(cell,),
            bounding_box=BoundingBox(1, 1, 1, 1),
            shape=(1, 1),
            range_a1="S!A1",
        )
        payload = InputGroupsPayload(
            workbook_name="wb.xlsx",
            generated_at_utc="2026-01-01T00:00:00Z",
            summary=InputGroupsSummary(
                total_groups=1,
                total_cells=1,
                orientation_histogram={"rowwise": 1},
            ),
            groups=(group,),
        )
        restored = InputGroupsPayload.from_dict(payload.to_dict())
        assert restored.summary.total_groups == 1
        assert restored.groups[0].group_id == "s_a1"
        assert restored.groups[0].cells[0].address == normalize_key("S!A1")


class TestSetterGenerationOptions:
    def test_label_mode_ignored_when_labels_disabled(self) -> None:
        opts = SetterGenerationOptions(include_labels=False, label_mode="all")
        assert opts.effective_label_mode() == "none"
