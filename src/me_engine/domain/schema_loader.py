"""Load market schema JSON files produced by schema extraction."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping


@dataclass(frozen=True, slots=True)
class DimensionSchema:
    title: str
    segments: tuple[str, ...]
    parent: Mapping[str, str | None] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class GeographyNodeSchema:
    name: str
    children: tuple["GeographyNodeSchema", ...] = ()


@dataclass(frozen=True, slots=True)
class MarketSchema:
    market_name: str
    dimensions: tuple[DimensionSchema, ...]
    geographies: tuple[str, ...]
    priced_dimension: str
    geography_tree: GeographyNodeSchema | None = None


def _parse_geography_node(raw: Mapping[str, Any]) -> GeographyNodeSchema:
    children = tuple(
        _parse_geography_node(c) for c in raw.get("children", [])
    )
    return GeographyNodeSchema(name=str(raw["name"]), children=children)


def _parse_dimension(raw: Mapping[str, Any]) -> DimensionSchema:
    parent_raw = raw.get("parent") or {}
    parent = {str(k): (None if v is None else str(v)) for k, v in parent_raw.items()}
    return DimensionSchema(
        title=str(raw["title"]),
        segments=tuple(str(s) for s in raw["segments"]),
        parent=parent,
    )


def load_schema(schema_path: Path | str) -> MarketSchema:
    """Load and parse a market schema JSON file."""
    path = Path(schema_path)
    with path.open(encoding="utf-8") as fh:
        raw = json.load(fh)

    tree_raw = raw.get("geography_tree")
    tree = _parse_geography_node(tree_raw) if tree_raw else None

    return MarketSchema(
        market_name=str(raw["market_name"]),
        dimensions=tuple(_parse_dimension(d) for d in raw["dimensions"]),
        geographies=tuple(str(g) for g in raw["geographies"]),
        priced_dimension=str(raw["priced_dimension"]),
        geography_tree=tree,
    )


def schema_to_dict(schema: MarketSchema) -> dict[str, Any]:
    """Serialize a MarketSchema back to JSON-compatible dict."""

    def _node_to_dict(node: GeographyNodeSchema) -> dict[str, Any]:
        return {
            "name": node.name,
            "children": [_node_to_dict(c) for c in node.children],
        }

    return {
        "market_name": schema.market_name,
        "dimensions": [
            {
                "title": d.title,
                "segments": list(d.segments),
                **({"parent": dict(d.parent)} if d.parent else {}),
            }
            for d in schema.dimensions
        ],
        "geographies": list(schema.geographies),
        "priced_dimension": schema.priced_dimension,
        **({"geography_tree": _node_to_dict(schema.geography_tree)}
           if schema.geography_tree else {}),
    }


def save_schema(schema: MarketSchema, schema_path: Path | str) -> Path:
    """Write a MarketSchema to disk."""
    path = Path(schema_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(schema_to_dict(schema), fh, indent=2, ensure_ascii=False)
        fh.write("\n")
    return path
