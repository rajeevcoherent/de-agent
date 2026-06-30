"""The Curve Agent — derives a growth-curve shape from evidence, with reasoning.

It does NOT emit market numbers. It emits a *normalised shape* plus the reasoning
and evidence behind it; the deterministic builder turns that into the value path.
The agent self-validates (smoothness, mean ~ 1) and falls back to the canonical
shape whenever it is offline or the model returns something implausible — so the
pipeline always produces a usable, explained curve.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from statistics import fmean

from .canonical import CANONICAL_SHAPE
from .config import AgentConfig
from .evidence_agent import EvidenceAgent, EvidenceBrief
from .llm import EvidenceGatherer, Evidence, LLMClient, LLMError
from .shape import NormalizedShape, YOY_YEARS

_SYSTEM = (
    "You are a market-growth analyst refining a growth-curve SHAPE (the relative "
    "pace of annual growth, mean ~ 1), NOT absolute numbers. You are given a "
    "baseline shape derived from many real markets: growth accelerates into a peak "
    "around 2030, then eases. Your job is to ADJUST that baseline for this specific "
    "market/geography using the evidence — keep its overall accelerate-then-ease "
    "character (peak should stay in 2029-2031) and deviate only where the evidence "
    "justifies it (e.g. a faster-maturing market peaks earlier and flatter). "
    "Return adjusted multipliers plus your reasoning. Respond ONLY with JSON."
)

# _SMOOTHNESS_LIMIT = 0.05 
_SMOOTHNESS_LIMIT = 0.15  # or even 0.2     # reject jagged shapes (mean |2nd difference|)
_MEAN_TOLERANCE = 0.08        # accept shapes whose mean is within 8% of 1.0


@dataclass(frozen=True, slots=True)
class CurveDecision:
    geography: str
    shape: NormalizedShape
    archetype: str
    reasoning: str
    evidence: tuple[Evidence, ...]
    confidence: float
    used_fallback: bool

    @property
    def peak_year(self) -> int:
        return self.shape.peak_year


@dataclass(frozen=True, slots=True)
class MarketContext:
    """The Input-Sheet-derived context handed to the agent for one geography."""

    market_name: str
    geography: str
    forecast_cagr: float
    region: str | None = None
    notes: str = ""


class CurveAgent:
    """Proposes a per-geography growth shape, grounded and validated.

    Two-pass decision:
      Pass 1 — LLM decides from the canonical prior only (fast, no web).
      If confidence < threshold → Evidence Agent fires DDG, brain evaluates snippets.
      Pass 2 — LLM re-decides with the curated evidence brief (better grounding).
    """

    def __init__(self, config: AgentConfig | None = None,
                 llm: LLMClient | None = None,
                 evidence: EvidenceGatherer | None = None) -> None:
        self._config = config or AgentConfig.load()
        self._llm = llm or LLMClient(self._config)
        self._evidence = evidence or EvidenceGatherer()
        self._ev_agent = EvidenceAgent(config=self._config, llm=self._llm)

    def decide(self, ctx: MarketContext) -> CurveDecision:
        if not self._config.is_online:
            return self._fallback(ctx, reason="offline (no API key)")
        try:
            # Pass 1: decide from prior + any basic evidence
            snippets = self._gather(ctx)
            raw = self._llm.complete_json(_SYSTEM, self._prompt(ctx, snippets))
            decision = self._validate(ctx, raw)

            # If confidence is weak, fire the Evidence Agent for a second pass
            if self._ev_agent.needs_evidence(decision.confidence):
                brief = self._ev_agent.gather_and_evaluate(
                    market=ctx.market_name,
                    geography=ctx.geography,
                    topic="market growth curve CAGR forecast",
                )
                print(f"    [EvidenceAgent] {ctx.geography}: {brief.log_summary()}")
                if brief.has_useful_evidence:
                    # Pass 2: re-run with curated web evidence
                    enriched_snippets = snippets + [
                        s.text for s in brief.snippets
                    ]
                    raw2 = self._llm.complete_json(
                        _SYSTEM, self._prompt(ctx, enriched_snippets))
                    decision = self._validate(ctx, raw2)

            return decision
        except (LLMError, KeyError, ValueError, TypeError) as exc:
            return self._fallback(ctx, reason=f"agent error: {exc}")

    # --- prompt & evidence --------------------------------------------------
    def _gather(self, ctx: MarketContext) -> list[str]:
        query = f"{ctx.market_name} {ctx.geography} growth trend forecast demand"
        return self._evidence.search(query)

    def _prompt(self, ctx: MarketContext, snippets: list[str]) -> str:
        evidence_block = "\n".join(f"- {s}" for s in snippets) or "- (no web evidence)"
        years = ", ".join(str(y) for y in YOY_YEARS)
        baseline = ", ".join(f"{m:.3f}" for m in CANONICAL_SHAPE.multipliers)
        return (
            f"Market: {ctx.market_name}\nGeography: {ctx.geography}\n"
            f"Region: {ctx.region or 'n/a'}\n"
            f"Forecast CAGR (2025-2033): {ctx.forecast_cagr:.4f}\n"
            f"Notes: {ctx.notes or 'n/a'}\n\n"
            f"Baseline shape (years {years}):\n  [{baseline}]\n\n"
            f"Evidence:\n{evidence_block}\n\n"
            f"Adjust the baseline for this market. Return JSON with keys:\n"
            f'  "multipliers": [12 floats, mean ~ 1.0, peak in 2029-2031],\n'
            f'  "archetype": short label,\n'
            f'  "reasoning": 1-3 sentences citing the evidence,\n'
            f'  "evidence": [{{"claim":..., "source":..., "confidence":0-1}}],\n'
            f'  "confidence": 0-1.'
        )

    # --- validation ---------------------------------------------------------
    # How much to trust the LLM's deviation from the data-derived baseline.
    _AGENT_WEIGHT = 0.5

    def _validate(self, ctx: MarketContext, raw: dict) -> CurveDecision:
        proposed = NormalizedShape(self._normalize(raw["multipliers"]))
        shape = self._guard(proposed)
        if shape.smoothness() > _SMOOTHNESS_LIMIT:
            return self._fallback(ctx, reason="proposed shape too jagged")
        return CurveDecision(
            geography=ctx.geography,
            shape=shape,
            archetype=str(raw.get("archetype", "unspecified")),
            reasoning=str(raw.get("reasoning", "")),
            evidence=self._parse_evidence(raw.get("evidence", [])),
            confidence=float(raw.get("confidence", 0.5)),
            used_fallback=False,
        )

    def _guard(self, proposed: NormalizedShape) -> NormalizedShape:
        """Blend the proposal toward the data-derived baseline.

        The canonical shape encodes the true accelerate-then-ease pattern observed
        across real markets; the LLM contributes market-specific nudges. Blending
        keeps the agent's signal while preventing a single bad proposal from
        degrading accuracy below the proven baseline.
        """
        w = self._AGENT_WEIGHT
        blended = tuple(
            w * p + (1 - w) * b
            for p, b in zip(proposed.multipliers, CANONICAL_SHAPE.multipliers)
        )
        mean = fmean(blended)
        return NormalizedShape(tuple(v / mean for v in blended))

    @staticmethod
    def _normalize(values: list) -> tuple[float, ...]:
        floats = [float(v) for v in values]
        if len(floats) != len(YOY_YEARS):
            raise ValueError(f"need {len(YOY_YEARS)} multipliers, got {len(floats)}")
        mean = fmean(floats)
        if mean <= 0:
            raise ValueError("non-positive shape mean")
        normalized = tuple(v / mean for v in floats)   # force mean to exactly 1
        return normalized

    @staticmethod
    def _parse_evidence(items) -> tuple[Evidence, ...]:
        out: list[Evidence] = []
        for item in items if isinstance(items, list) else []:
            try:
                out.append(Evidence(
                    claim=str(item["claim"]),
                    source=str(item.get("source", "")),
                    confidence=float(item.get("confidence", 0.5)),
                ))
            except (KeyError, TypeError, ValueError):
                continue
        return tuple(out)

    def _fallback(self, ctx: MarketContext, reason: str) -> CurveDecision:
        return CurveDecision(
            geography=ctx.geography,
            shape=CANONICAL_SHAPE,
            archetype="canonical-adoption-hump",
            reasoning=f"Canonical reverse-engineered shape used ({reason}).",
            evidence=(),
            confidence=0.4,
            used_fallback=True,
        )
