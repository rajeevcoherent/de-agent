"""Helpers to turn the generation agent into a runnable fleet step."""
from __future__ import annotations

from pathlib import Path

from .generation_agent import GenerationAgent


def generate_input_sheet(market_name: str, run_pipeline: bool = False) -> Path:
    agent = GenerationAgent()
    plan = agent.plan(market_name)
    output_path = agent.materialize(market_name)
    if run_pipeline:
        from generate_input import run_pipeline

        output_me = Path(f"my_outputs/ME-{market_name}.xlsx")
        report_path = Path(f"output/report-{market_name}.json")
        run_pipeline(output_path, output_me, None, report_path, False)
    return output_path
