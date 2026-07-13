"""Batch runner for all Input Sheet workbooks in a folder.

This script discovers Excel workbooks under an input folder and invokes
run_pipeline.py for each one, writing generated ME workbooks and JSON reports.

Example:
    python run_all.py --input-folder my_inputs --out-dir output --report-dir output/reports
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run run_pipeline.py over all input workbooks")
    p.add_argument("--input-folder", default="my_inputs", help="Folder containing input .xlsx workbooks")
    p.add_argument("--out-dir", default="my_outputs", help="Folder to write generated ME workbooks")
    p.add_argument("--report-dir", default="output", help="Folder to write JSON report files")
    p.add_argument("--reference-folder", default="my_reference",
                   help="Folder containing reference ME workbooks for Gate 3 diff")
    p.add_argument("--template-folder", default=None,
                   help="Optional folder containing template workbooks for styling")
    p.add_argument("--save-metadata", action="store_true",
                   help="Save workbook inspector metadata for each input file")
    p.add_argument("--no-agents", action="store_true",
                   help="Skip agents and use flat priors only for all files")
    p.add_argument("--fail-fast", action="store_true",
                   help="Stop processing on the first failing input file")
    return p.parse_args()


def _find_workbooks(folder: Path) -> list[Path]:
    if not folder.exists():
        raise FileNotFoundError(f"Input folder does not exist: {folder}")
    return sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx"])


def _normalize_name(path: Path) -> str:
    name = path.stem.lower()
    for remove in ["data file -", "data file-", "data file", "input sheet -", "input sheet-", "input sheet", "me -", "me-", "me"]:
        if name.startswith(remove):
            name = name[len(remove):].strip()
            break
    name = name.replace("_", " ").replace("-", " ").replace("&", " ")
    normalized = " ".join(word for word in name.split() if word)
    return normalized


def _market_name_from_input(path: Path) -> str:
    normalized = _normalize_name(path)
    return " ".join(word.capitalize() if word not in {"and", "of", "by", "for", "the"} else word for word in normalized.split())


def _find_reference_path(reference_folder: Path, input_path: Path) -> Path | None:
    normalized_input = _normalize_name(input_path)
    if not reference_folder.exists():
        return None

    best_match = None
    for candidate in sorted(reference_folder.iterdir()):
        if not candidate.is_file() or candidate.suffix.lower() != ".xlsx":
            continue
        normalized_reference = _normalize_name(candidate)
        if normalized_reference == normalized_input:
            return candidate
        if normalized_reference in normalized_input or normalized_input in normalized_reference:
            best_match = candidate
    return best_match


def _find_template_path(template_folder: Path, input_path: Path) -> Path | None:
    normalized_input = _normalize_name(input_path)
    if not template_folder.exists():
        return None

    best_match = None
    for candidate in sorted(template_folder.iterdir()):
        if not candidate.is_file() or candidate.suffix.lower() != ".xlsx":
            continue
        normalized_template = _normalize_name(candidate)
        if normalized_template == normalized_input:
            return candidate
        if normalized_template in normalized_input or normalized_input in normalized_template:
            best_match = candidate
    return best_match


def _run_pipeline_for_file(
    run_pipeline_path: Path,
    input_path: Path,
    output_path: Path,
    report_path: Path,
    reference_path: Path | None,
    template_path: Path | None,
    save_metadata: bool,
    no_agents: bool,
) -> tuple[int, str]:
    args = [sys.executable, str(run_pipeline_path), "--input", str(input_path), "--output", str(output_path)]
    if reference_path is not None:
        args.extend(["--reference", str(reference_path)])
    if template_path is not None:
        args.extend(["--template", str(template_path)])
    if save_metadata:
        args.append("--save-metadata")
    if report_path is not None:
        args.extend(["--save-report", str(report_path)])
    if no_agents:
        args.append("--no-agents")

    proc = subprocess.run(args, capture_output=True, text=True)
    return proc.returncode, proc.stdout + proc.stderr


def main() -> int:
    args = _parse_args()
    input_folder = Path(args.input_folder)
    out_dir = Path(args.out_dir)
    report_dir = Path(args.report_dir)
    reference_folder = Path(args.reference_folder) if args.reference_folder else None
    template_folder = Path(args.template_folder) if args.template_folder else None
    run_pipeline_path = Path(__file__).parent / "run_pipeline.py"

    out_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)

    workbooks = _find_workbooks(input_folder)
    if not workbooks:
        print(f"No .xlsx workbooks found in {input_folder}")
        return 1

    summary: list[dict[str, object]] = []
    print(f"Found {len(workbooks)} workbook(s) in {input_folder}")

    for input_path in workbooks:
        market_name = _market_name_from_input(input_path)
        output_path = out_dir / f"ME-{market_name}.xlsx"
        report_path = report_dir / f"report-{market_name}.json"
        reference_path = None
        template_path = None

        if reference_folder is not None:
            reference_path = _find_reference_path(reference_folder, input_path)
            if reference_path is None:
                print(f"  [WARN] Reference not found for {input_path.name} in {reference_folder}")
        if template_folder is not None:
            template_path = _find_template_path(template_folder, input_path)
            if template_path is None:
                print(f"  [WARN] Template not found for {input_path.name} in {template_folder}")

        print(f"\nRunning pipeline for: {input_path.name}")
        returncode, output = _run_pipeline_for_file(
            run_pipeline_path,
            input_path,
            output_path,
            report_path,
            reference_path,
            template_path,
            save_metadata=args.save_metadata,
            no_agents=args.no_agents,
        )

        summary.append({
            "input": str(input_path),
            "output": str(output_path),
            "report": str(report_path),
            "reference": str(reference_path) if reference_path is not None else None,
            "returncode": returncode,
        })

        print(output)
        if returncode != 0:
            print(f"  [ERROR] Pipeline failed for {input_path.name} (exit {returncode})")
            if args.fail_fast:
                break

    summary_path = out_dir / "batch_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(f"\nBatch complete. Summary written to: {summary_path}")

    failed = [row for row in summary if row["returncode"] != 0]
    if failed:
        print(f"{len(failed)} workbook(s) failed")
        return 1

    print("All workbooks processed successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
