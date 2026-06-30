"""Build runtime taxonomy objects — either from a schema file or directly
from an Input Sheet workbook (the preferred, schema-free path).

The pipeline uses ``build_taxonomy_from_workbook()`` which inspects the workbook
in memory and returns a ``RuntimeTaxonomy`` without touching any JSON files.
``load_runtime_taxonomy()`` is kept for tooling and backward compatibility.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .schema_loader import GeographyNodeSchema, MarketSchema, load_schema
from .taxonomy import Dimension, Geography, GeographyIndex, _g


@dataclass(frozen=True, slots=True)
class RuntimeTaxonomy:
    """Market-specific taxonomy built entirely in memory."""

    market_name: str
    geographies: GeographyIndex
    segmentation_dimensions: tuple[Dimension, ...]
    priced_dimension: Dimension

    @property
    def GEOGRAPHIES(self) -> GeographyIndex:
        return self.geographies

    @property
    def SEGMENTATION_DIMENSIONS(self) -> tuple[Dimension, ...]:
        return self.segmentation_dimensions

    @property
    def PRICED_DIMENSION(self) -> Dimension:
        return self.priced_dimension


# ---------------------------------------------------------------------------
# Build helpers shared by both paths
# ---------------------------------------------------------------------------

def _build_geography_tree(node: GeographyNodeSchema) -> Geography:
    return _g(node.name, *(_build_geography_tree(c) for c in node.children))


def _flat_tree_from_leaves(leaves: tuple[str, ...]) -> Geography:
    """Fallback: Global root with all geographies as direct children."""
    return _g("Global", *(_g(name) for name in leaves))


def build_runtime_taxonomy(schema: MarketSchema) -> RuntimeTaxonomy:
    """Convert a MarketSchema into runtime Dimension / Geography objects."""
    dimensions = tuple(
        Dimension(d.title, d.segments, parent=d.parent)
        for d in schema.dimensions
    )
    priced = next(d for d in dimensions if d.title == schema.priced_dimension)

    if schema.geography_tree:
        root = _build_geography_tree(schema.geography_tree)
    else:
        root = _flat_tree_from_leaves(schema.geographies)

    return RuntimeTaxonomy(
        market_name=schema.market_name,
        geographies=GeographyIndex(root),
        segmentation_dimensions=dimensions,
        priced_dimension=priced,
    )


def load_runtime_taxonomy(schema_path: Path | str) -> RuntimeTaxonomy:
    """Load schema JSON and build runtime taxonomy (for tooling / compat)."""
    return build_runtime_taxonomy(load_schema(schema_path))


# ---------------------------------------------------------------------------
# Schema-free path: build directly from the Input Sheet workbook
# ---------------------------------------------------------------------------

def build_taxonomy_from_workbook(input_path: Path | str) -> RuntimeTaxonomy:
    """Inspect an Input Sheet workbook and return a RuntimeTaxonomy in memory.

    This is the primary pipeline entry point.  No JSON file is written or read.
    The WorkbookInspector extraction logic is reused via extract_schema_in_memory()
    which calls the same routines but returns a MarketSchema object directly.
    """
    from me_engine.tools.extract_schema import extract_schema
    schema = extract_schema(input_path)
    return build_runtime_taxonomy(schema)


def default_taxonomy() -> RuntimeTaxonomy:
    """Return the module-default taxonomy (olive oil schema)."""
    from . import taxonomy as _tax

    return RuntimeTaxonomy(
        market_name="",
        geographies=_tax.GEOGRAPHIES,
        segmentation_dimensions=_tax.SEGMENTATION_DIMENSIONS,
        priced_dimension=_tax.PRICED_DIMENSION,
    )
