"""Immutable taxonomy of the market-estimation model.

Generic infrastructure types (years, bands, dimension/geography structures)
live here. Market-specific segments and geographies are loaded at runtime from
JSON schemas via ``runtime_taxonomy.load_runtime_taxonomy``.

For backward compatibility, module-level ``GEOGRAPHIES``, ``PRICED_DIMENSION``,
and ``SEGMENTATION_DIMENSIONS`` are populated from ``schemas/olive_oil.json``.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from functools import cached_property
from pathlib import Path
from typing import Mapping

# --- Time horizon -----------------------------------------------------------
BASE_YEAR = 2025          # the anchor year (Value here == the global anchor)
FIRST_YEAR = 2021         # historical start (back-cast)
LAST_YEAR = 2033          # forecast horizon
YEARS: tuple[int, ...] = tuple(range(FIRST_YEAR, LAST_YEAR + 1))


class Band(str, Enum):
    """The three vertically stacked bands present in every geography sheet."""

    VALUE = "Market Value (US$ Mn)"
    ASP = "ASP"
    VOLUME = "Market Volume (Th Liters)"


@dataclass(frozen=True, slots=True)
class Dimension:
    """A segmentation dimension (e.g. 'By Product Type') and its ordered members."""

    title: str
    segments: tuple[str, ...]
    parent: Mapping[str, str | None] = field(default_factory=dict)

    def parent_of(self, segment: str) -> str | None:
        return self.parent.get(segment)


@dataclass(frozen=True, slots=True)
class Geography:
    """A node in the geography tree (Global, a region, or a country)."""

    name: str
    children: tuple["Geography", ...] = ()

    @property
    def is_leaf(self) -> bool:
        return not self.children


def _g(name: str, *children: Geography) -> Geography:
    return Geography(name, tuple(children))


@dataclass(frozen=True)
class GeographyIndex:
    """Flattened, queryable view of the geography tree."""

    root: Geography

    @cached_property
    def in_order(self) -> tuple[Geography, ...]:
        """All geographies, parents before children (sheet order)."""
        return tuple(self._walk(self.root))

    @cached_property
    def by_name(self) -> Mapping[str, Geography]:
        return {g.name: g for g in self.in_order}

    @cached_property
    def parent_of(self) -> Mapping[str, str | None]:
        parents: dict[str, str | None] = {self.root.name: None}
        for parent in self.in_order:
            for child in parent.children:
                parents[child.name] = parent.name
        return parents

    @staticmethod
    def _walk(node: Geography):
        yield node
        for child in node.children:
            yield from GeographyIndex._walk(child)


def _build_geo_node(node) -> Geography:
    return _g(node.name, *(_build_geo_node(c) for c in node.children))


def _load_legacy_defaults():
    """Load olive-oil schema for backward-compatible module-level exports."""
    from .schema_loader import load_schema

    schema_path = Path(__file__).resolve().parents[3] / "schemas" / "olive_oil.json"
    schema = load_schema(schema_path)
    dimensions = tuple(
        Dimension(d.title, d.segments, parent=d.parent) for d in schema.dimensions
    )
    priced = next(d for d in dimensions if d.title == schema.priced_dimension)
    root = (_build_geo_node(schema.geography_tree) if schema.geography_tree
            else _g("Global"))
    return GeographyIndex(root), dimensions, priced


GEOGRAPHIES, SEGMENTATION_DIMENSIONS, PRICED_DIMENSION = _load_legacy_defaults()
