"""Market-agnostic ME pipeline runner — fully runtime-driven.

Usage:
    python run_pipeline.py --input  "Data File - XYZ Market.xlsx" \\
                           --output "my_outputs/ME - XYZ Market.xlsx" \\
                           [--reference "my_reference/ME - XYZ Market.xlsx"] \\
                           [--template  "ME - Global Avocado Oil Market.xlsx"] \\
                           [--no-agents]

Flow:
    Input Workbook
        │
        ▼
    WorkbookInspector  (extract structure, optional debug JSON)
        │
        ▼
    RuntimeTaxonomy    (built in memory — NO schemas/*.json read or written)
        │
        ├── Gate 1    (validate input against RuntimeTaxonomy)
        ├── Agents    (Curve + Segmentation + ASP)
        ├── Assembler (pure math)
        └── Validator (Gate 3 identity + reference diff)
        │
        ▼
    Output ME Workbook → Reference Comparison
"""
from __future__ import annotations

import argparse
import io
import json
import sys
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

# Windows console: force UTF-8 so box-drawing chars don't crash
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent / "src"))

from me_engine.assembly.assembler import Assembler
from me_engine.domain.runtime_taxonomy import build_taxonomy_from_workbook
from me_engine.io.input_drivers import InputDriverBuilder
from me_engine.io.me_writer import MEWorkbookWriter
from me_engine.io.workbook_inspector import WorkbookInspector, metadata_output_path
from me_engine.validation.gate1 import InputValidator
from me_engine.validation.gate3 import OutputValidator


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="ME Pipeline Runner")
    p.add_argument("--input",     required=True,  help="Path to Input Sheet .xlsx")
    p.add_argument("--output",    required=True,  help="Path to write the generated ME workbook")
    p.add_argument("--reference", default=None,   help="Optional: ground-truth ME workbook for Gate 3 diff")
    p.add_argument("--template",  default=None,   help="Style template workbook (defaults to reference if provided)")
    p.add_argument("--no-agents", action="store_true", help="Skip agents, use flat priors only")
    p.add_argument("--save-metadata", action="store_true",
                   help="Write WorkbookInspector JSON to output/extracted_metadata/ (debug only)")
    p.add_argument("--save-report", default=None,
                   help="Write JSON summary report after pipeline completes")
    return p.parse_args()


def _collect_agent_rationale(drivers) -> dict[str, str]:
    """Collect per-geography curve agent reasoning from the builder."""
    if hasattr(drivers, "curve_rationale"):
        return drivers.curve_rationale()
    return {}


def main() -> int:
    args = _parse_args()
    input_path  = Path(args.input)
    output_path = Path(args.output)
    ref_path    = Path(args.reference) if args.reference else None
    template    = Path(args.template) if args.template else ref_path
    use_agents  = not args.no_agents

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # ── WORKBOOK INSPECTION ─────────────────────────────────────────────────
    print("\n┌─ WORKBOOK INSPECTOR ───────────────────────────────────────────")
    meta = WorkbookInspector().inspect(input_path)
    for line in meta.summary_lines():
        print(line)
    if not meta.passed():
        # Inspection FAIL means the workbook is unreadable or missing required
        # sheets — halt immediately, nothing else can run.
        print("  [FAIL] Workbook inspection failed — pipeline cannot continue.")
        for issue in meta.issues:
            if issue.level == "FAIL":
                print(f"    {issue}")
        if args.save_metadata:
            meta_path = metadata_output_path(input_path, output_path.parent / "extracted_metadata")
            meta.save(meta_path)
            print(f"  Debug metadata → {meta_path}")
        print("└─ PIPELINE HALTED.\n")
        return 1

    if args.save_metadata:
        meta_path = metadata_output_path(input_path, output_path.parent / "extracted_metadata")
        meta.save(meta_path)
        print(f"  Debug metadata → {meta_path}")
    print("└─ Inspection done.\n")

    # ── RUNTIME TAXONOMY (in memory, no JSON files) ─────────────────────────
    print("┌─ RUNTIME TAXONOMY ─────────────────────────────────────────────")
    taxonomy = build_taxonomy_from_workbook(input_path)
    print(f"  Market   : {taxonomy.market_name}")
    print(f"  Seg dims : {len(taxonomy.segmentation_dimensions)}")
    print(f"  Geos     : {len(taxonomy.geographies.by_name)}")
    print(f"  Priced   : {taxonomy.priced_dimension.title} "
          f"({len(taxonomy.priced_dimension.segments)} products)")
    print("└─ Taxonomy built.\n")

    # ── GATE 1 ──────────────────────────────────────────────────────────────
    print("┌─ GATE 1: Input Validation ─────────────────────────────────────")
    gate1 = InputValidator(taxonomy).validate(input_path)
    print(gate1)
    if not gate1.passed():
        print("└─ PIPELINE HALTED — fix input errors above before proceeding.\n")
        return 1
    print("└─ Gate 1 passed.\n")

    # ── AGENTS + GATE 2 ─────────────────────────────────────────────────────
    print("┌─ AGENTS: Curve + Segmentation + ASP ───────────────────────────")
    truth_me = str(ref_path) if ref_path else None
    builder = InputDriverBuilder(
        input_path,
        truth_me=truth_me,
        use_segmentation_agent=use_agents,
        taxonomy=taxonomy,
    )
    drivers = builder.build()
    print(f"  Assembled DriverSet for {len(drivers.geographies)} geographies.")
    print("└─ Agents done (Gate 2 constraints enforced per-agent).\n")

    # ── ASSEMBLER ───────────────────────────────────────────────────────────
    print("┌─ ASSEMBLER: Building MarketResult ─────────────────────────────")
    result = Assembler(taxonomy).assemble(drivers)
    print(f"  Built {len(result.geographies)} geography results.")
    print("└─ Assembly done.\n")

    # ── WRITE OUTPUT ────────────────────────────────────────────────────────
    if template and template.exists():
        writer = MEWorkbookWriter(template)
    elif ref_path and ref_path.exists():
        writer = MEWorkbookWriter(ref_path)
    else:
        fallback = Path("ME - Global Avocado Oil Market.xlsx")
        if not fallback.exists():
            print("ERROR: no template workbook found. Provide --template or --reference.")
            return 1
        print(f"  [WARN] Using avocado template for styling: {fallback}")
        writer = MEWorkbookWriter(fallback)

    writer.write(result, output_path)
    print(f"  Wrote: {output_path}\n")

    # ── GATE 3 ──────────────────────────────────────────────────────────────
    print("┌─ GATE 3: Output Validation ─────────────────────────────────────")
    rationale = _collect_agent_rationale(builder)
    gate3 = OutputValidator(taxonomy).validate(result, ref_path, rationale)
    print(gate3)
    if gate3.passed():
        print("└─ Gate 3 passed.\n")
    else:
        print("└─ Gate 3 FAILED — see violations above.\n")

    if args.save_report:
        report_path = Path(args.save_report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_data = {
            "market": result.market_name,
            "output": str(output_path.resolve()),
            "reference": str(ref_path.resolve()) if ref_path else None,
            "gate1_passed": gate1.passed(),
            "gate3_identity_passed": gate3.identity.passed(),
            "gate3_passed": gate3.passed(),
            "cells_compared": gate3.diff.cells_compared if gate3.diff else 0,
            "mean_rel_error": gate3.diff.mean_rel_error if gate3.diff else None,
            "median_rel_error": gate3.diff.median_rel_error if gate3.diff else None,
            "p90_rel_error": gate3.diff.p90_rel_error if gate3.diff else None,
            "worst_rel_error": gate3.diff.worst_rel_error if gate3.diff else None,
            "cells_within_5pct": gate3.diff.cells_within_5pct if gate3.diff else None,
        }
        report_path.write_text(json.dumps(report_data, indent=2), encoding="utf-8")
        print(f"  Wrote report: {report_path}\n")

    # ── SUMMARY ─────────────────────────────────────────────────────────────
    print("=" * 65)
    print(f"  Market  : {result.market_name}")
    print(f"  Geos    : {len(result.geographies)}")
    print(f"  Output  : {output_path}")
    if gate3.diff:
        print(f"  vs Ref  : {gate3.diff.mean_rel_error:.2%} mean error, "
              f"{gate3.diff.cells_within_5pct:.1%} within 5%")
    print("=" * 65)

    return 0 if gate3.passed() else 1


if __name__ == "__main__":
    sys.exit(main())
