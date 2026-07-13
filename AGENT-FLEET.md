# The Agent Fleet — Sourcing Every Driver From Evidence

The deterministic assembler needs a complete `DriverSet`. The agent fleet produces
it from the Input Sheet + evidence, one specialist per driver family. Every agent
emits **judgement + reasoning**, never final market numbers — the assembler does
all arithmetic. All agents run on a cheap model (`openai/gpt-4o-mini` via
OpenRouter), cache responses on disk, and fall back to a data-derived prior when
offline or on bad output.

## Design principle: data-anchored agents
Each driver was reverse-engineered from the ME file first, yielding a **prior**
(canonical curve, segment-premium directions, ~0.85% ASP creep). Agents *adjust*
the prior with market-specific reasoning rather than inventing from scratch. This
keeps accuracy at the proven baseline while adding explainability — a single bad
LLM proposal cannot drop the result below the prior.

## The agents

### 1. Curve Agent  (`curve/agent.py`)
- **Decides:** the normalised year-by-year growth shape per geography.
- **Prior:** the canonical accelerate-to-2030-then-ease shape (cosine 0.9997 to truth).
- **Adjusts:** peak timing/steepness from GDP, adoption stage, maturity evidence.
- **Guardrail:** blended 50/50 with the prior; peak constrained to 2029-2031.
- **Result (live, 37 geos):** mean path error **2.34%**, cosine 0.9997, 0 fallbacks.

### 2. Segmentation Agent  (`seg_agent.py` + `segmentation.py`)
- **Decides:** each segment's annual growth premium vs the market (who gains share).
- **Math:** premium -> per-segment value path -> re-normalised drifting shares
  (reproduces ME drift to **0.118%** when given true premiums).
- **Reality:** LLM gets share-drift *direction* right for major segments, magnitudes
  approximate; clamped to +/-3%/yr.

### 3. ASP Agent  (`asp_agent.py`)
- **Decides:** per-product annual price-inflation rate.
- **Prior:** data-derived **0.85%/yr** (ASP is near-uniform across products/geos).
- **Math:** compound from the base-year price to a full ASP path; volume = value/asp.

### 4. Generation Agent  (`generation_agent.py`)
- **Decides:** the market-specific blueprint for an Input Sheet from local schema/template evidence.
- **Input:** a market description or market name.
- **Output:** a structured plan with schema hint, output workbook name/path, and default CAGR/anchor values.
- **Materialization:** can create a workbook in the agent-generated folder and hand it off to the existing pipeline.

## Infrastructure
- **LLM client** (`llm.py`): OpenAI-compatible JSON calls, provider-agnostic.
- **Cache** (`cache.py`): SQLite content-addressed cache — re-runs are free.
- **Evidence** (`llm.py`): best-effort DuckDuckGo snippets to ground reasoning.
- **Config** (`config.py`): keys/model from `.env`; offline => canonical fallback.

## Accuracy harness
Because the real ME file exists, every agent decision is scored against ground
truth (per-geography curve score; per-cell value deviation). See PILOT-RESULTS.md
for the end-to-end numbers.

## Known gaps (next work)
- Segmentation magnitudes need a multi-market premium prior (like the curve has).
- Global/region sheets are bottom-up aggregations not yet generated from countries.
- The discovered `CMI_Trend_NN` tags (in the Input Sheet) are ground truth for
  *which* trend each segment used — a future check on agent trend selection.
