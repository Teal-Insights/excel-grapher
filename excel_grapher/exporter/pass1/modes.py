"""Shared types for internals refactor and formula clustering."""

from __future__ import annotations

from typing import Literal, TypeAlias, cast, get_args

VariationMode: TypeAlias = Literal["independent", "dominant_key_only"]
ClusteringMode: TypeAlias = Literal["series", "series_ast", "ast"]


def unwrap_annotation(annotation: object) -> object:
    """Resolve typing aliases that expose ``__value__`` (PEP 695 / TypeAliasType).

    ``typing.get_args`` does not always evaluate those wrappers, so callers that
    introspect ``Literal`` aliases at runtime must unwrap first. Plain
    ``TypeAlias`` annotations need no unwrap on Python 3.11+.
    """
    while True:
        value = getattr(annotation, "__value__", None)
        if value is None or value is annotation:
            return annotation
        annotation = value


def _literal_args(alias: object) -> tuple[object, ...]:
    return get_args(unwrap_annotation(alias))


_VARIATION_MODE_CHOICES: tuple[str, ...] = cast(tuple[str, ...], _literal_args(VariationMode))
_CLUSTERING_MODE_CHOICES: tuple[str, ...] = cast(tuple[str, ...], _literal_args(ClusteringMode))

VARIATION_MODE_CLI_HELP = (
    "Formula-cluster variation mode for internals refactor clustering "
    '(default: workbook_config.VARIATION_MODE or "independent"). '
    "Only affects export and refactor-bucket stages, not --extract-graph."
)

CLUSTERING_MODE_CLI_HELP = (
    "Formula-cluster base mode for internals refactor clustering "
    '(default: workbook_config.CLUSTERING_MODE or "series_ast"). '
    "Only affects export and refactor-bucket stages, not --extract-graph."
)


def variation_mode_choices() -> tuple[str, ...]:
    return _VARIATION_MODE_CHOICES


def clustering_mode_choices() -> tuple[str, ...]:
    return _CLUSTERING_MODE_CHOICES


def parse_variation_mode(value: object) -> VariationMode:
    if value not in _VARIATION_MODE_CHOICES:
        raise ValueError(
            f"VARIATION_MODE must be one of {list(_VARIATION_MODE_CHOICES)}; got {value!r}"
        )
    return cast(VariationMode, value)


def parse_clustering_mode(value: object) -> ClusteringMode:
    if value not in _CLUSTERING_MODE_CHOICES:
        raise ValueError(
            f"CLUSTERING_MODE must be one of {list(_CLUSTERING_MODE_CHOICES)}; got {value!r}"
        )
    return cast(ClusteringMode, value)
