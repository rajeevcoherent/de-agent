**Project: ME pipeline — Automation & Status**

Purpose
- Ingest Input Sheet workbooks (`.xlsx`), extract structural metadata, run agents and assembly to produce an ME workbook, then compare the generated ME to a reference and emit comparison metrics.

What the pipeline does (exact behavior)
- Inspect Input Sheet with `WorkbookInspector` (dimensions, geographies, ASP, validation issues).
- Build an in-memory `RuntimeTaxonomy` from the workbook.
- Gate 1: validate input against the taxonomy (halt/mark if failing).
- Run agents (Curve, Segmentation, ASP) and apply per-agent constraints.
- Assemble the ME workbook and write it to the configured `--output` path.
- Gate 3: compare generated ME with provided reference ME and compute metrics.

Produced artifacts (per input file)
- `output/<slug>_ME.xlsx` — generated ME workbook
- `output/extracted_metadata/<slug>_metadata.json` — Inspector metadata (if `--save-metadata` used)
- `output/reports/<slug>_report.json` — comparison report (see keys below)
- `output/logs/<timestamp>.log` — run-level summary and warnings (optional)

Comparison/report JSON keys (written by `--save-report`)
- `market`, `output`, `reference`, `gate1_passed`, `gate3_passed`
- `cells_compared`, `mean_rel_error`, `median_rel_error`, `p90_rel_error`, `cells_within_5pct`, `worst_rel_error`

Recent robustness change (prevents pipeline break on new sheets)
- `src/me_engine/domain/runtime_taxonomy.py`: `build_taxonomy_from_workbook()` now attempts strict schema extraction and, on exception, falls back to `WorkbookInspector()` to construct a best-effort in-memory `MarketSchema` so the pipeline can continue. The inspector records `WARN`/`FAIL` issues for review. See [src/me_engine/domain/runtime_taxonomy.py](src/me_engine/domain/runtime_taxonomy.py).

Failure & retry policy (recommended and implemented)
- Try strict extraction once. On failure, infer schema in-memory and continue (implemented).
- If Gate 1 still fails, mark `<slug>_needs_review.json` with issues and continue processing other files.
- Retry once for transient I/O/parsing errors, then mark `needs_review`.

Automation (Planned) — single-entrypoint runner (concrete)
- CLI: `run_all.py --input-folder my_inputs/ --out-dir output/ --report-dir output/reports/ --workers N`
- Behavior:
  - discover `.xlsx` files in `--input-folder`
  - for each file: run `run_pipeline.py --input <file> --output <out_dir>/<slug>_ME.xlsx --save-metadata --save-report <report_dir>/<slug>_report.json`
  - on failure: follow retry policy, write `<slug>_needs_review.json` and continue
  - produce `output/summary_batch.csv` or JSON listing file statuses and key metrics

Example commands
- Single file (existing runner):

```
python run_pipeline.py --input "my_inputs/Data File - X.xlsx" --output "output/X_ME.xlsx" --save-metadata --save-report "output/reports/X_report.json"
```

- Bulk (proposed):

```
python run_all.py --input-folder my_inputs/ --out-dir output/ --report-dir output/reports/ --workers 4
```


Next steps
- Owner will implement automation, CI, and tests as required. This document is a concise status summary for managers; implementation details and timelines will be tracked separately.

If you want, I can export this document to PDF for sharing.
