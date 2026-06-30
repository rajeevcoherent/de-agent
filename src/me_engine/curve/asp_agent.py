"""The ASP Agent — decides price-growth (inflation) paths per product.

Reverse-engineering showed ASP rises ~0.8-0.9%/yr, near-uniform across products
and geographies (a mild price creep). So the agent's judgement is a small annual
ASP-inflation rate per product; the math compounds it from the base-year price.
The agent emits only the rate + rationale, never the price path itself.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping

from .config import AgentConfig
from .llm import LLMClient, LLMError
from ..domain.series import Series
from ..domain.taxonomy import BASE_YEAR, Dimension, YEARS

# Data-derived default (mean ASP CAGR across the avocado ME file).
DEFAULT_ASP_INFLATION = 0.0085
_MAX_ASP_INFLATION = 0.05

_SYSTEM = (
    "You are a pricing analyst. Estimate the annual average-selling-price growth "
    "(inflation) for each product in a market, as a small rate (typically 0.00 to "
    "0.03). Premium products may rise slightly faster. Respond ONLY with JSON."
)


@dataclass(frozen=True, slots=True)
class AspDecision:
    rates: Mapping[str, float]
    rationale: str
    used_fallback: bool


class AspAgent:
    """Chooses per-product ASP inflation rates and builds price paths."""

    def __init__(self, config: AgentConfig | None = None,
                 llm: LLMClient | None = None) -> None:
        self._config = config or AgentConfig.load()
        self._llm = llm or LLMClient(self._config)

    def decide(self, market_name: str,
               priced_dimension: Dimension | None = None) -> AspDecision:
        dim = priced_dimension
        if dim is None:
            from ..domain.taxonomy import PRICED_DIMENSION
            dim = PRICED_DIMENSION
        if not self._config.is_online:
            return self._fallback("offline", dim)
        try:
            raw = self._llm.complete_json(_SYSTEM, self._prompt(market_name, dim))
            return self._validate(raw, dim)
        except (LLMError, KeyError, ValueError, TypeError):
            return self._fallback("agent error", dim)

    def price_path(self, base_price: float, rate: float) -> Series:
        return Series({
            year: base_price * (1.0 + rate) ** (year - BASE_YEAR)
            for year in YEARS
        })

    def _prompt(self, market_name: str, dim: Dimension) -> str:
        products = "\n".join(f"  - {p}" for p in dim.segments)
        return (
            f"Market: {market_name}\nProducts:\n{products}\n\n"
            f'Return JSON: {{"rates": [{{"product": <name>, "rate": <float>}}], '
            f'"rationale": <short>}} covering every product.'
        )

    def _validate(self, raw: dict, dim: Dimension) -> AspDecision:
        provided = {r["product"]: r for r in raw.get("rates", [])
                    if isinstance(r, dict) and "product" in r}
        rates = {
            product: self._clamp(provided.get(product, {}).get("rate"))
            for product in dim.segments
        }
        return AspDecision(rates, str(raw.get("rationale", "")), used_fallback=False)

    @staticmethod
    def _clamp(value) -> float:
        try:
            rate = float(value)
        except (TypeError, ValueError):
            return DEFAULT_ASP_INFLATION
        return max(0.0, min(_MAX_ASP_INFLATION, rate))

    @staticmethod
    def _fallback(reason: str, dim: Dimension) -> AspDecision:
        rates = {p: DEFAULT_ASP_INFLATION for p in dim.segments}
        return AspDecision(rates, f"data-derived default ({reason})",
                           used_fallback=True)
