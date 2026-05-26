"""Shared validation/resolution issue helpers."""

from __future__ import annotations

from typing import Literal

from excel_grapher.series_bindings.types import ValidationIssue


def make_issue(
    level: Literal["error", "warning"],
    code: str,
    message: str,
    *,
    series_id: str | None = None,
    address: str | None = None,
) -> ValidationIssue:
    return {
        "level": level,
        "code": code,
        "message": message,
        "series_id": series_id,
        "address": address,
    }
