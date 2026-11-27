"""Command-line interface (`excel-parse`) for the Layer 1 `excel_parsing` SDK."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from excel_parsing.models import ParserConfig, ProgressEvent
from excel_parsing.pipeline import ExcelParser


def _print_progress(ev: ProgressEvent) -> None:
    if ev.status == "failed":
        print(
            f"  [{ev.completed_regions}/{ev.total_regions}] FAILED {ev.sheet_name} "
            f"R{ev.region_id:02d} ({ev.coord_range}): {ev.error}",
            file=sys.stderr,
        )
    else:
        acct = "OK" if ev.accounting_passed else "FAIL"
        print(
            f"  [{ev.completed_regions}/{ev.total_regions}] {ev.sheet_name} "
            f"R{ev.region_id:02d} ({ev.coord_range}): valid={ev.valid_items}, "
            f"issues={ev.validation_issues}, acct={acct} ({ev.elapsed_seconds}s)",
            file=sys.stderr,
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="excel-parse",
        description="Inspect or extract structured catalog items from Excel workbooks.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # excel-parse inspect
    p_ins = sub.add_parser("inspect", help="Run fast CPU-only Stages 1-2 inspection (< 1s).")
    p_ins.add_argument("workbook", type=Path, help="Path to `.xlsx` workbook.")
    p_ins.add_argument(
        "--sheet",
        action="append",
        dest="sheets",
        help="Worksheet name to inspect (repeatable; defaults to all visible sheets).",
    )
    p_ins.add_argument(
        "--include-html",
        action="store_true",
        help="Include compact HTML string in JSON output.",
    )

    # excel-parse extract
    p_ext = sub.add_parser("extract", help="Run full Stages 1-4 extraction + PostgreSQL export.")
    p_ext.add_argument("workbook", type=Path, help="Path to `.xlsx` workbook.")
    p_ext.add_argument(
        "--sheet",
        action="append",
        dest="sheets",
        help="Worksheet name to extract (repeatable; defaults to all visible sheets).",
    )
    p_ext.add_argument(
        "--vllm-url",
        default="http://localhost:8005/v1",
        help="vLLM OpenAI-compatible base URL (default: http://localhost:8005/v1).",
    )
    p_ext.add_argument(
        "--model",
        default=None,
        help="Explicit model ID (auto-detected from `/v1/models` if omitted).",
    )
    p_ext.add_argument(
        "--concurrency",
        type=int,
        default=6,
        help="Max concurrent vLLM requests (default: 6).",
    )
    p_ext.add_argument(
        "--pg-table",
        default="equipment_catalog",
        help="Target PostgreSQL table name for `json_to_recordset` output.",
    )
    p_ext.add_argument(
        "--param-style",
        choices=("dollar", "pyformat", "named"),
        default="dollar",
        help="SQL bind parameter style: dollar ($1::json), pyformat (%%s::json), or named (:records).",
    )
    p_ext.add_argument(
        "--on-conflict",
        default=None,
        help="Optional ON CONFLICT clause body for PostgreSQL INSERT.",
    )
    p_ext.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Directory to persist HTML previews, per-region JSONs, review Markdown, and `_pg_recordset.json`.",
    )
    p_ext.add_argument(
        "--fail-on-accounting",
        action="store_true",
        help="Exit non-zero if any sheet fails the Row-Accounting Invariant.",
    )
    p_ext.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress progress logs on stderr (emit only JSON payload on stdout).",
    )

    args = parser.parse_args(argv)

    if args.command == "inspect":
        sdk = ExcelParser()
        inspection = sdk.inspect(args.workbook, sheet_names=args.sheets)
        print(json.dumps(inspection.to_dict(include_html=args.include_html), ensure_ascii=False, indent=2))
        return 0

    cfg = ParserConfig(
        vllm_base_url=args.vllm_url,
        vllm_model=args.model,
        max_concurrent_requests=args.concurrency,
        max_region_workers=args.concurrency,
    )
    sdk = ExcelParser(config=cfg)
    result = sdk.extract(
        args.workbook,
        sheet_names=args.sheets,
        fail_on_accounting=args.fail_on_accounting,
        output_dir=args.output_dir,
        pg_table=args.pg_table,
        on_progress=None if args.quiet else _print_progress,
    )

    pg_payload = result.to_pg_recordset(
        table=args.pg_table,
        param_style=args.param_style,
        on_conflict=args.on_conflict,
    )
    print(pg_payload.to_json(indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
