"""Market-analysis agent that plans and materializes an Input Sheet."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
import re
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from me_engine.curve.config import AgentConfig
from me_engine.curve.llm import EvidenceGatherer, LLMClient, LLMError
from me_engine.domain.schema_loader import DimensionSchema, GeographyNodeSchema, MarketSchema, load_schema
from me_engine.tools.extract_schema import extract_schema


@dataclass(frozen=True, slots=True)
class GenerationPlan:
    market_name: str
    schema_hint: str | None
    output_name: str
    output_path: str
    cagr: float
    anchor: float


@dataclass(frozen=True, slots=True)
class EvidenceBranch:
    label: str
    parent: str | None
    confidence: float


class GenerationAgent:
    """Create a structured generation plan for a market from local templates or web-grounded reasoning."""

    def __init__(self, root: Path | None = None, llm: LLMClient | None = None, evidence: EvidenceGatherer | None = None, config: AgentConfig | None = None) -> None:
        base = Path(root) if root is not None else Path(__file__).resolve().parents[3]
        self._root = base.resolve()
        self._schemas_dir = self._root / "schemas"
        self._inputs_dir = self._root / "my_inputs"
        self._outputs_dir = self._root / "agentGeneratedInputSheet"
        self._config = config or AgentConfig.load(self._root)
        self._llm = llm or LLMClient(self._config)
        self._evidence = evidence or EvidenceGatherer()

    def plan(self, market_name: str) -> GenerationPlan:
        schema_hint = self._infer_schema_hint(market_name)
        output_name = f"Data File - {market_name}.xlsx"
        output_path = str((self._outputs_dir / output_name).resolve())
        return GenerationPlan(
            market_name=market_name,
            schema_hint=schema_hint,
            output_name=output_name,
            output_path=output_path,
            cagr=0.08,
            anchor=100.0,
        )

    def materialize(self, market_name: str) -> Path:
        plan = self.plan(market_name)
        self._outputs_dir.mkdir(parents=True, exist_ok=True)
        output_path = Path(plan.output_path)
        if output_path.exists():
            return output_path

        schema = self._load_schema(plan.schema_hint, market_name)
        from generate_input import create_input_workbook

        return create_input_workbook(schema, market_name, output_path, plan.cagr, plan.anchor)

    def _load_schema(self, schema_hint: str | None, market_name: str) -> MarketSchema:
        schema_path = self._resolve_schema_path(schema_hint, market_name)
        if schema_path is not None:
            if schema_path.suffix.lower() == ".xlsx":
                return extract_schema(schema_path)
            try:
                return load_schema(schema_path)
            except KeyError:
                return self._schema_from_json(schema_path, market_name)

        return self._plan_schema_from_web(market_name)

    def _schema_from_json(self, schema_path: Path, market_name: str) -> MarketSchema:
        from me_engine.domain.schema_loader import DimensionSchema

        raw = json.loads(schema_path.read_text(encoding="utf-8"))
        dimensions = []
        if "product_type" in raw:
            dimensions.append(DimensionSchema(title="By Product Type", segments=tuple(raw["product_type"])))
        if "dimensions" in raw:
            for entry in raw["dimensions"]:
                dimensions.append(DimensionSchema(
                    title=str(entry.get("title", "By Segment")),
                    segments=tuple(str(seg) for seg in entry.get("segments", [])),
                    parent={str(k): (None if v is None else str(v)) for k, v in (entry.get("parent") or {}).items()},
                ))
        if not dimensions:
            raise ValueError("no dimensions discovered")
        return MarketSchema(
            market_name=str(raw.get("market_name", market_name)),
            dimensions=tuple(dimensions),
            geographies=tuple(str(g) for g in raw.get("geographies", [])),
            priced_dimension=raw.get("priced_dimension", dimensions[0].title if dimensions else "By Product Type"),
            geography_tree=None,
        )

    def _plan_schema_from_web(self, market_name: str) -> MarketSchema:
        snippets = self._evidence.search(self._web_query(market_name))
        evidence_block = "\n".join(f"- {s}" for s in snippets) if snippets else "- (no web evidence)"
        if not snippets:
            return self._fallback_schema(market_name)

        raw = self._plan_from_evidence(market_name, snippets)
        dimensions = []
        for entry in raw.get("dimensions", []) or []:
            title = str(entry.get("title", "By Segment"))
            segments = tuple(str(seg) for seg in entry.get("segments", []) or [])
            parent = {str(k): (None if v is None else str(v)) for k, v in (entry.get("parent") or {}).items()}
            if segments:
                dimensions.append(DimensionSchema(title=title, segments=segments, parent=parent))
        if not dimensions:
            return self._fallback_schema(market_name)

        geographies = tuple(str(g) for g in raw.get("geographies", []) or self._default_geographies(market_name))
        tree_raw = raw.get("geography_tree")
        geography_tree = self._parse_geography_tree(tree_raw) if isinstance(tree_raw, dict) else None
        return MarketSchema(
            market_name=str(raw.get("market_name", market_name)),
            dimensions=tuple(dimensions),
            geographies=geographies,
            priced_dimension=str(raw.get("priced_dimension", dimensions[0].title if dimensions else "By Product Type")),
            geography_tree=geography_tree,
        )

    def _plan_from_evidence(self, market_name: str, snippets: list[str]) -> dict:
        evidence_block = "\n".join(f"- {s}" for s in snippets)
        system = (
            "You are a market-structure planner. Given a market name and multiple web snippets, "
            "extract a schema. Return JSON with keys: market_name, dimensions, geographies, "
            "priced_dimension, geography_tree, and an optional evidence_summary."
        )
        user = (
            f"Market: {market_name}\n"
            f"Evidence:\n{evidence_block}\n\n"
            "For each candidate segment dimension, infer a small set of parent/child branches. "
            "Only keep a deeper branch if the evidence strongly supports it. Return a compact schema."
        )
        try:
            raw = self._llm.complete_json(system, user)
        except (LLMError, KeyError, TypeError, ValueError):
            return {}

        if not isinstance(raw, dict):
            return {}

        normalized_dimensions = []
        for entry in raw.get("dimensions", []) or []:
            title = str(entry.get("title", "By Segment"))
            segments = [str(seg) for seg in entry.get("segments", []) or []]
            parent = {str(k): (None if v is None else str(v)) for k, v in (entry.get("parent") or {}).items()}
            if not segments:
                continue
            branches = self._evidence_branches(snippets, segments, parent)
            kept_segments = []
            kept_parent = {}
            for branch in branches:
                if branch.confidence < 0.55:
                    continue
                kept_segments.append(branch.label)
                if branch.parent is not None:
                    kept_parent[branch.label] = branch.parent
            if kept_segments:
                normalized_dimensions.append({
                    "title": title,
                    "segments": kept_segments,
                    "parent": kept_parent,
                })
        if not normalized_dimensions:
            return {}

        normalized_geographies = []
        for geo in raw.get("geographies", []) or []:
            normalized_geographies.append(str(geo))
        if not normalized_geographies:
            normalized_geographies = self._default_geographies(market_name)

        tree = raw.get("geography_tree")
        if tree is not None and isinstance(tree, dict):
            tree = self._prune_tree(tree, snippets)

        return {
            "market_name": raw.get("market_name", market_name),
            "dimensions": normalized_dimensions,
            "geographies": normalized_geographies,
            "priced_dimension": raw.get("priced_dimension", "By Product Type"),
            "geography_tree": tree,
        }

    def _evidence_branches(self, snippets: list[str], segments: list[str], parent: dict) -> list[EvidenceBranch]:
        text = "\n".join(snippets).lower()
        branches: list[EvidenceBranch] = []
        for seg in segments:
            label = str(seg)
            parent_label = parent.get(label)
            score = 0.35
            if label.lower() in text:
                score += 0.25
            if parent_label and parent_label.lower() in text:
                score += 0.2
            if any(keyword in text for keyword in ("market", "segment", "category", "subsegment", "country", "region")):
                score += 0.1
            if self._looks_like_deeper_branch(label, parent_label):
                score -= 0.1
            score = min(0.95, max(0.1, score))
            branches.append(EvidenceBranch(label=label, parent=parent_label, confidence=round(score, 2)))
        return branches

    @staticmethod
    def _looks_like_deeper_branch(label: str, parent: str | None) -> bool:
        lowered = label.lower()
        if parent is None:
            return False
        return any(token in lowered for token in ("premium", "value", "pro", "plus", "advanced", "smart", "ultra"))

    def _prune_tree(self, tree: dict, snippets: list[str]) -> dict:
        text = "\n".join(snippets).lower()
        if not isinstance(tree, dict):
            return tree

        def prune(node: dict) -> dict | None:
            name = str(node.get("name", ""))
            children = []
            for child in node.get("children", []) or []:
                if not isinstance(child, dict):
                    continue
                child_name = str(child.get("name", "")).lower()
                if child_name and child_name in text:
                    pruned = prune(child)
                    if pruned is not None:
                        children.append(pruned)
                elif name.lower() in text or not children:
                    pruned = prune(child)
                    if pruned is not None:
                        children.append(pruned)
            if not children and name.lower() not in text and name.lower() not in {"world", "north america", "europe", "asia pacific", "latin america", "middle east", "africa", "u.s.", "canada", "germany", "u.k.", "japan", "china"}:
                return None
            return {"name": name, "children": children}

        pruned = prune(tree)
        return pruned if isinstance(pruned, dict) else {"name": "World", "children": []}

    def _fallback_schema(self, market_name: str) -> MarketSchema:
        if "supplement" in market_name.lower() or "protein" in market_name.lower():
            dimensions = (
                DimensionSchema(
                    title="By Product Type",
                    segments=("Protein Supplements", "Whey Protein", "Plant Protein", "Collagen Protein"),
                    parent={"Whey Protein": "Protein Supplements", "Plant Protein": "Protein Supplements", "Collagen Protein": "Protein Supplements"},
                ),
                DimensionSchema(
                    title="By Consumer",
                    segments=("Consumers", "Athletes", "General Wellness", "Seniors"),
                    parent={"Athletes": "Consumers", "General Wellness": "Consumers", "Seniors": "Consumers"},
                ),
            )
            geographies = ("U.S.", "Canada", "Germany", "U.K.")
            geography_tree = GeographyNodeSchema(
                name="World",
                children=(
                    GeographyNodeSchema(name="North America", children=(GeographyNodeSchema(name="U.S."), GeographyNodeSchema(name="Canada"))),
                    GeographyNodeSchema(name="Europe", children=(GeographyNodeSchema(name="Germany"), GeographyNodeSchema(name="U.K."))),
                ),
            )
        else:
            dimensions = (
                DimensionSchema(title="By Product Type", segments=("Core", "Premium", "Value")),
                DimensionSchema(title="By Channel", segments=("Direct", "Retail", "Online")),
            )
            geographies = ("U.S.", "Germany", "Japan", "China")
            geography_tree = None
        return MarketSchema(
            market_name=market_name,
            dimensions=dimensions,
            geographies=geographies,
            priced_dimension="By Product Type",
            geography_tree=geography_tree,
        )

    def _default_geographies(self, market_name: str) -> tuple[str, ...]:
        if "supplement" in market_name.lower() or "protein" in market_name.lower():
            return ("U.S.", "Canada", "Germany", "U.K.")
        return ("U.S.", "Germany", "Japan", "China")

    def _parse_geography_tree(self, raw: dict | None) -> GeographyNodeSchema | None:
        if not isinstance(raw, dict):
            return None

        def parse(node: dict) -> GeographyNodeSchema:
            children = tuple(parse(child) for child in node.get("children", []) if isinstance(child, dict))
            return GeographyNodeSchema(name=str(node.get("name", "World")), children=children)

        return parse(raw)

    def _web_query(self, market_name: str) -> str:
        return f"{market_name} market segmentation regions countries demand growth"

    def _infer_schema_hint(self, market_name: str) -> str | None:
        normalized = self._slugify(market_name)
        for path in sorted(self._schemas_dir.glob("*.json")):
            if path.stem == normalized:
                return str(path)
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                continue
            if str(payload.get("market_name", "")).strip().lower() == market_name.strip().lower():
                return str(path)
        return None

    def _resolve_schema_path(self, schema_hint: str | None, market_name: str) -> Path | None:
        if schema_hint:
            candidate = Path(schema_hint)
            if candidate.exists():
                return candidate

        normalized = self._slugify(market_name)
        for path in sorted(self._schemas_dir.glob("*.json")):
            if path.stem == normalized:
                return path

        for path in sorted(self._schemas_dir.glob("*.json")):
            if normalized in self._slugify(path.stem) or self._slugify(path.stem) in normalized:
                return path

        if self._inputs_dir.exists():
            for path in sorted(self._inputs_dir.glob("*.xlsx")):
                path_slug = self._slugify(path.stem)
                if normalized in path_slug or path_slug in normalized:
                    return path

        if self._inputs_dir.exists():
            for path in sorted(self._inputs_dir.glob("*.xlsx")):
                if "adhesive" in normalized or "sealant" in normalized:
                    if "adhesive" in self._slugify(path.stem) or "sealant" in self._slugify(path.stem):
                        return path

        return None

    @staticmethod
    def _slugify(value: str) -> str:
        normalized = value.strip().lower()
        for ch in [" ", "-", "/", "&", ".", ","]:
            normalized = normalized.replace(ch, "_")
        normalized = "_".join(part for part in normalized.split("_") if part)
        return normalized
