"""Shared Pass-1 synthesis models (issue #595)."""

from __future__ import annotations

from dataclasses import dataclass

from excel_grapher.exporter.pass1.bindings import BindingKeyValue


@dataclass(frozen=True)
class MemberContext:
    """One formula-cell member of a Pass-1 cluster or singleton unit."""

    address: str
    function_name: str
    engine_column: str
    normalized_formula: str
    python_source: str
    dependency_addresses: tuple[str, ...]
    dependency_functions: tuple[str, ...]
    binding_keys: dict[str, BindingKeyValue] | None = None
    binding_record: dict[str, BindingKeyValue] | None = None


class SeriesHelperVerificationError(ValueError):
    """Raised when a bound series cannot be verified into a Pass-1 helper."""
