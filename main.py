"""Example local runner demonstrating the `excel_parsing` Layer 1 Python SDK."""

from pathlib import Path

from excel_parsing import ExcelParser, ParserConfig, ProgressEvent


def _log_progress(ev: ProgressEvent) -> None:
    if ev.status == "failed":
        print(
            f"  FAILED Region #{ev.region_id} ({ev.coord_range}): {ev.error}"
        )
    else:
        acct = "OK" if ev.accounting_passed else "FAIL"
        print(
            f"  [{ev.completed_regions}/{ev.total_regions}] {ev.sheet_name} "
            f"Region #{ev.region_id} ({ev.coord_range}): "
            f"{ev.valid_items} valid, {ev.validation_issues} issues, acct={acct} | "
            f"{ev.elapsed_seconds}s"
        )


if __name__ == "__main__":
    excel_path = Path(
        "/mnt/workspace/projects/excel_parsing/samples/catalogs/●ツメ表ひな多島【面数換算】_1～20島_20260108.xlsx"
    )
    output_dir = Path("/mnt/workspace/projects/excel_parsing/outputs")

    parser = ExcelParser(
        ParserConfig(
            vllm_base_url="http://localhost:8005/v1",
            max_concurrent_requests=6,
        )
    )

    # Stage 1-2 fast inspection (< 1s, zero LLM calls)
    inspection = parser.inspect(excel_path)
    for s in inspection.sheets:
        print(
            f"Inspect Sheet: {s.sheet_name} | visible_cells={s.visible_cells} "
            f"(hidden_cols={s.hidden_cols}) | regions={len(s.regions)} "
            f"(extractable={s.extractable_region_count}) | html={len(s.compact_html):,} chars"
        )

    # Stage 3-4 full extraction + reconciliation + PostgreSQL json_to_recordset export
    result = parser.extract(
        excel_path,
        output_dir=output_dir,
        pg_table="equipment_catalog",
        on_progress=_log_progress,
    )

    pg_payload = result.to_pg_recordset(table="equipment_catalog")
    print(
        f"\nWorkbook complete in {result.elapsed_seconds}s | "
        f"all_accounting_passed={result.all_accounting_passed} | "
        f"total_valid_items={pg_payload.row_count}"
    )
    print(f"PostgreSQL query: {pg_payload.query}")
