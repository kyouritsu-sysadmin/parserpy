"""Stateless Layer 1 SDK pipeline: `ExcelParser`, `inspect_workbook`, and `extract_workbook`."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, replace
import hashlib
import io
import json
from pathlib import Path
import time
from typing import Literal
import zipfile

import openpyxl
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_parsing.exceptions import (
    InvalidWorkbookError,
    RowAccountingError,
    SheetNotFoundError,
)
from excel_parsing.extraction.extract import (
    ExtractionResult,
    VLLMClient,
    extract_region,
    save_extraction_result,
)
from excel_parsing.models import (
    ExcelSource,
    ParserConfig,
    PreprocessedSheet,
    ProgressEvent,
    RegionSummary,
    SheetExtractionResult,
    SheetInspection,
    WorkbookExtractionResult,
    WorkbookInspection,
)
from excel_parsing.preprocess.decompose import resolve_merged_cells
from excel_parsing.preprocess.detection import (
    Region,
    detect_regions,
    save_regions_to_excel,
)
from excel_parsing.preprocess.formatting import (
    classify_region_rows,
    extract_formatting,
    save_formatted_regions_to_excel,
)
from excel_parsing.preprocess.normalize import ErrorStrategy, normalize_grid
from excel_parsing.reconcile.reconcile import (
    RegionReconciliation,
    assemble_sheet_reconciliation,
    reconcile_region,
    save_reconciliation,
)
from excel_parsing.serialize.serialize import save_html, sheet_to_html


def is_extractable_region(region: Region) -> bool:
    """True when `region` represents an extractable data table or item strip."""
    if region.kind == "vertical_strip":
        return region.num_rows >= 3
    if region.kind == "table":
        return not (region.num_rows <= 2 and region.num_cols <= 2 and region.cell_count <= 3)
    return False


def _read_workbook_bytes(source: ExcelSource) -> bytes:
    if isinstance(source, bytes):
        raw = source
    elif isinstance(source, bytearray):
        raw = bytes(source)
    elif isinstance(source, (str, Path)):
        path = Path(source)
        if not path.is_file():
            raise InvalidWorkbookError(f"Workbook file not found: {path}")
        raw = path.read_bytes()
    elif hasattr(source, "read"):
        data = source.read()
        if isinstance(data, str):
            raise InvalidWorkbookError("Binary stream required for `.xlsx` source, got text stream")
        raw = bytes(data)
    else:
        raise InvalidWorkbookError(
            f"Unsupported Excel source type: {type(source).__name__}. "
            "Expected str, Path, bytes, bytearray, or BinaryIO."
        )

    if len(raw) < 4 or raw[:4] != b"PK\x03\x04":
        raise InvalidWorkbookError("Input is not a valid `.xlsx` (ZIP/OOXML) archive")
    return raw


def _open_workbooks(raw_bytes: bytes, *, load_formulas: bool = False) -> tuple[Workbook, Workbook | None]:
    try:
        wb_values = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        wb_formulas = (
            openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=False)
            if load_formulas
            else None
        )
        return wb_values, wb_formulas
    except (zipfile.BadZipFile, KeyError, ValueError, OSError) as exc:
        raise InvalidWorkbookError(f"Failed to load `.xlsx` workbook: {exc}") from exc


def _select_worksheets(
    wb_values: Workbook,
    sheet_names: Sequence[str] | None,
) -> list[Worksheet]:
    if sheet_names is not None:
        selected: list[Worksheet] = []
        for name in sheet_names:
            if name not in wb_values.sheetnames:
                raise SheetNotFoundError(
                    f"Worksheet {name!r} not found. Available sheets: {wb_values.sheetnames}"
                )
            selected.append(wb_values[name])
        return selected

    visible = [ws for ws in wb_values.worksheets if ws.sheet_state == "visible"]
    if visible:
        return visible
    return [wb_values.active] if wb_values.active is not None else []


def preprocess_sheet(
    ws: Worksheet,
    *,
    formula_ws: Worksheet | None = None,
    skip_hidden: bool = True,
    connectivity: Literal[4, 8] = 4,
    row_gap: int = 1,
    col_gap: int = 0,
    merge_overlapping_bboxes: bool = False,
    split_banner_bridges: bool = True,
    split_side_by_side: bool = True,
    reunite_split_columns: bool = True,
    min_cells: int = 1,
    error_strategy: ErrorStrategy = "tag",
    normalize_unicode: bool = True,
    collapse_spaces: bool = True,
    format_dates_iso: bool = True,
) -> PreprocessedSheet:
    """Run Stage 1 spatial decomposition & preprocessing on a single worksheet."""
    raw_grid = resolve_merged_cells(ws, skip_hidden=skip_hidden)

    norm_grid, report = normalize_grid(
        raw_grid,
        error_strategy=error_strategy,
        normalize_unicode=normalize_unicode,
        collapse_spaces=collapse_spaces,
        format_dates_iso=format_dates_iso,
        formula_ws=formula_ws,
    )

    regions = detect_regions(
        norm_grid,
        connectivity=connectivity,
        row_gap=row_gap,
        col_gap=col_gap,
        merge_overlapping_bboxes=merge_overlapping_bboxes,
        split_banner_bridges=split_banner_bridges,
        split_side_by_side=split_side_by_side,
        reunite_split_columns=reunite_split_columns,
        min_cells=min_cells,
    )

    formats = extract_formatting(ws, norm_grid)
    row_roles = {
        region.region_id: classify_region_rows(region, formats)
        for region in regions
    }

    return PreprocessedSheet(
        sheet_name=ws.title,
        grid=norm_grid,
        regions=regions,
        formats=formats,
        row_roles=row_roles,
        report=report,
    )


def preprocess_workbook(
    source: ExcelSource,
    *,
    sheet_names: Sequence[str] | None = None,
    load_formulas: bool = False,
    **sheet_kwargs,
) -> list[PreprocessedSheet]:
    """Load a workbook from path, bytes, or stream and preprocess selected sheets."""
    raw_bytes = _read_workbook_bytes(source)
    wb_values, wb_formulas = _open_workbooks(raw_bytes, load_formulas=load_formulas)
    target_sheets = _select_worksheets(wb_values, sheet_names)

    results: list[PreprocessedSheet] = []
    for ws in target_sheets:
        formula_ws = wb_formulas[ws.title] if wb_formulas is not None else None
        results.append(preprocess_sheet(ws, formula_ws=formula_ws, **sheet_kwargs))
    return results


class ExcelParser:
    """Reusable SDK client for workbook inspection, LLM extraction, and PostgreSQL export.

    Holds a shared `VLLMClient` (keep-alive HTTP pool + concurrency semaphore) so
    backend services and workers can safely call `inspect()` and `extract()` across
    many requests.
    """

    def __init__(
        self,
        config: ParserConfig | None = None,
        *,
        client: VLLMClient | None = None,
        **config_overrides,
    ) -> None:
        base_cfg = config or ParserConfig()
        self.config = replace(base_cfg, **config_overrides) if config_overrides else base_cfg
        self._client = client or VLLMClient(
            base_url=self.config.vllm_base_url,
            model=self.config.vllm_model,
            api_key=self.config.vllm_api_key,
            timeout=self.config.vllm_timeout,
            default_max_tokens=self.config.default_max_tokens,
            disable_thinking=self.config.disable_thinking,
            max_retries=self.config.max_retries,
            max_concurrent_requests=self.config.max_concurrent_requests,
        )

    @property
    def client(self) -> VLLMClient:
        return self._client

    def _preprocess_from_bytes(
        self,
        raw_bytes: bytes,
        sheet_names: Sequence[str] | None = None,
    ) -> tuple[list[str], list[str], list[PreprocessedSheet]]:
        cfg = self.config
        wb_values, wb_formulas = _open_workbooks(raw_bytes, load_formulas=cfg.load_formulas)
        all_names = list(wb_values.sheetnames)
        vis_names = [ws.title for ws in wb_values.worksheets if ws.sheet_state == "visible"]
        target_sheets = _select_worksheets(wb_values, sheet_names)

        preprocessed: list[PreprocessedSheet] = []
        for ws in target_sheets:
            formula_ws = wb_formulas[ws.title] if wb_formulas is not None else None
            preprocessed.append(
                preprocess_sheet(
                    ws,
                    formula_ws=formula_ws,
                    skip_hidden=cfg.skip_hidden,
                    connectivity=cfg.connectivity,
                    row_gap=cfg.row_gap,
                    col_gap=cfg.col_gap,
                    merge_overlapping_bboxes=cfg.merge_overlapping_bboxes,
                    split_banner_bridges=cfg.split_banner_bridges,
                    split_side_by_side=cfg.split_side_by_side,
                    reunite_split_columns=cfg.reunite_split_columns,
                    min_cells=cfg.min_cells,
                    error_strategy=cfg.error_strategy,
                    normalize_unicode=cfg.normalize_unicode,
                    collapse_spaces=cfg.collapse_spaces,
                    format_dates_iso=cfg.format_dates_iso,
                )
            )
        return all_names, vis_names, preprocessed

    def inspect(
        self,
        source: ExcelSource,
        *,
        sheet_names: Sequence[str] | None = None,
    ) -> WorkbookInspection:
        """Run fast CPU-only Stages 1-2 (< 1s) without making any LLM calls."""
        t0 = time.perf_counter()
        raw_bytes = _read_workbook_bytes(source)
        all_names, vis_names, sheets = self._preprocess_from_bytes(raw_bytes, sheet_names)

        inspections: list[SheetInspection] = []
        for s in sheets:
            compact_html = sheet_to_html(
                s.sheet_name,
                s.regions,
                s.formats,
                s.row_roles,
                include_coords=self.config.include_coords,
            )
            reg_summaries = [
                RegionSummary(
                    region_id=r.region_id,
                    coord_range=r.coord_range,
                    kind=r.kind,
                    rows=r.num_rows,
                    cols=r.num_cols,
                    cell_count=r.cell_count,
                    density=r.density,
                    extractable=is_extractable_region(r),
                )
                for r in s.regions
            ]
            inspections.append(
                SheetInspection(
                    sheet_name=s.sheet_name,
                    max_row=s.grid.max_row,
                    max_col=s.grid.max_col,
                    visible_cells=len(s.grid.cells),
                    merges=len(s.grid.merges),
                    hidden_rows=len(s.grid.hidden_rows),
                    hidden_cols=len(s.grid.hidden_cols),
                    errors=s.report.error_summary(),
                    blank_strings_removed=s.report.blank_strings_removed,
                    regions=reg_summaries,
                    compact_html=compact_html,
                )
            )

        return WorkbookInspection(
            all_sheet_names=all_names,
            visible_sheet_names=vis_names,
            sheets=inspections,
            elapsed_seconds=round(time.perf_counter() - t0, 4),
        )

    def extract(
        self,
        source: ExcelSource,
        *,
        sheet_names: Sequence[str] | None = None,
        fail_on_accounting: bool = False,
        output_dir: str | Path | None = None,
        pg_table: str = "equipment_catalog",
        save_debug_excel: bool = False,
        on_progress: Callable[[ProgressEvent], None] | None = None,
    ) -> WorkbookExtractionResult:
        """Run full Stages 1-4 extraction + reconciliation on `source`."""
        t_wb = time.perf_counter()
        raw_bytes = _read_workbook_bytes(source)
        content_sha256 = hashlib.sha256(raw_bytes).hexdigest()

        _, _, preprocessed_sheets = self._preprocess_from_bytes(raw_bytes, sheet_names)
        if preprocessed_sheets:
            self._client.resolve_model()

        out_base = Path(output_dir) if output_dir is not None else None
        cfg = self.config
        sheet_results: list[SheetExtractionResult] = []

        for sheet in preprocessed_sheets:
            t_sheet = time.perf_counter()
            compact_html = sheet_to_html(
                sheet.sheet_name,
                sheet.regions,
                sheet.formats,
                sheet.row_roles,
                include_coords=cfg.include_coords,
            )

            saved_files: dict[str, Path] = {}
            extraction_dir: Path | None = None
            if out_base is not None:
                out_base.mkdir(parents=True, exist_ok=True)
                saved_files["compact_html"] = save_html(
                    compact_html,
                    out_base / f"{sheet.sheet_name}_compact.html",
                    standalone=False,
                )
                saved_files["preview_html"] = save_html(
                    compact_html,
                    out_base / f"{sheet.sheet_name}_preview.html",
                    standalone=True,
                    title=sheet.sheet_name,
                )
                extraction_dir = out_base / "extraction" / sheet.sheet_name
                extraction_dir.mkdir(parents=True, exist_ok=True)

            summary_entries: list[dict] = []
            table_regions: list[Region] = []
            for region in sheet.regions:
                if not is_extractable_region(region):
                    summary_entries.append(
                        {
                            "region_id": region.region_id,
                            "coord_range": region.coord_range,
                            "kind": region.kind,
                            "status": "skipped_non_table",
                            "cell_count": region.cell_count,
                        }
                    )
                else:
                    table_regions.append(region)

            def _run_region(reg: Region):
                res = extract_region(
                    self._client,
                    reg,
                    sheet.formats,
                    sheet.row_roles.get(reg.region_id),
                    max_rows_per_chunk=cfg.max_rows_per_chunk,
                    include_coords=cfg.include_coords,
                    max_chunk_workers=cfg.max_chunk_workers,
                )
                rec = reconcile_region(
                    res,
                    region=reg,
                    client=self._client,
                    formats=sheet.formats,
                    row_roles=sheet.row_roles.get(reg.region_id),
                    retry_missing=cfg.retry_missing,
                )
                reg_saved = (
                    save_extraction_result(res, extraction_dir)
                    if extraction_dir is not None
                    else {}
                )
                return reg, res, reg_saved, rec

            region_extractions: list[ExtractionResult] = []
            region_reconciliations: list[RegionReconciliation] = []
            workers = max(1, min(cfg.max_region_workers, len(table_regions))) if table_regions else 1
            completed_cnt = 0

            with ThreadPoolExecutor(max_workers=workers) as pool:
                future_to_reg = {pool.submit(_run_region, reg): reg for reg in table_regions}
                for fut in as_completed(future_to_reg):
                    reg = future_to_reg[fut]
                    completed_cnt += 1
                    try:
                        region, result, reg_saved, rec = fut.result()
                    except Exception as exc:
                        err_msg = f"{type(exc).__name__}: {exc}"
                        summary_entries.append(
                            {
                                "region_id": reg.region_id,
                                "coord_range": reg.coord_range,
                                "kind": reg.kind,
                                "status": "failed",
                                "error": err_msg,
                            }
                        )
                        if on_progress is not None:
                            on_progress(
                                ProgressEvent(
                                    sheet_name=sheet.sheet_name,
                                    completed_regions=completed_cnt,
                                    total_regions=len(table_regions),
                                    region_id=reg.region_id,
                                    coord_range=reg.coord_range,
                                    status="failed",
                                    error=err_msg,
                                )
                            )
                        if fail_on_accounting:
                            raise
                        continue

                    region_extractions.append(result)
                    region_reconciliations.append(rec)
                    summary_entries.append(
                        {
                            "region_id": region.region_id,
                            "coord_range": region.coord_range,
                            "kind": region.kind,
                            "status": "extracted",
                            "skeleton_rows": len(result.skeleton),
                            "chunks": len(result.chunks),
                            "records": len(result.records),
                            "warnings": result.tree.warnings,
                            "accounting_passed": rec.accounting.passed,
                            "valid_items": len(rec.items),
                            "validation_issues": len(rec.validation_issues),
                            "retried_rows": rec.retried_rows,
                            "timing": asdict(result.timing),
                            "files": {k: str(v) for k, v in reg_saved.items()},
                        }
                    )
                    if on_progress is not None:
                        on_progress(
                            ProgressEvent(
                                sheet_name=sheet.sheet_name,
                                completed_regions=completed_cnt,
                                total_regions=len(table_regions),
                                region_id=region.region_id,
                                coord_range=region.coord_range,
                                status="extracted",
                                valid_items=len(rec.items),
                                validation_issues=len(rec.validation_issues),
                                accounting_passed=rec.accounting.passed,
                                elapsed_seconds=result.timing.total_seconds,
                            )
                        )

            region_extractions.sort(key=lambda r: r.region_id)
            summary_entries.sort(key=lambda e: e["region_id"])
            sheet_elapsed = round(time.perf_counter() - t_sheet, 2)

            sheet_reconciliation = assemble_sheet_reconciliation(
                sheet.sheet_name,
                region_reconciliations,
            )

            if extraction_dir is not None:
                summary_path = extraction_dir / "_summary.json"
                summary_path.write_text(
                    json.dumps(
                        {
                            "sheet_name": sheet.sheet_name,
                            "stage3_total_seconds": sheet_elapsed,
                            "regions": summary_entries,
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
                saved_files["summary"] = summary_path
                rec_files = save_reconciliation(
                    sheet_reconciliation,
                    extraction_dir,
                    pg_table=pg_table,
                )
                saved_files.update(rec_files)

            if out_base is not None and save_debug_excel:
                saved_files["detected_excel"] = save_regions_to_excel(
                    sheet.regions,
                    out_base / f"{sheet.sheet_name}_detected_regions.xlsx",
                    include_region_sheets=False,
                )
                saved_files["formatted_excel"] = save_formatted_regions_to_excel(
                    sheet.regions,
                    sheet.formats,
                    out_base / f"{sheet.sheet_name}_formatted_regions.xlsx",
                    include_region_sheets=False,
                )

            sheet_results.append(
                SheetExtractionResult(
                    sheet_name=sheet.sheet_name,
                    preprocessed=sheet,
                    compact_html=compact_html,
                    region_extractions=region_extractions,
                    reconciliation=sheet_reconciliation,
                    elapsed_seconds=sheet_elapsed,
                    saved_files=saved_files,
                )
            )

        wb_result = WorkbookExtractionResult(
            sheets=sheet_results,
            content_sha256=content_sha256,
            elapsed_seconds=round(time.perf_counter() - t_wb, 2),
        )

        if out_base is not None and sheet_results:
            wb_pg = wb_result.to_pg_recordset(table=pg_table)
            wb_pg.save(out_base / "extraction" / "_workbook_pg_recordset.json")

        if fail_on_accounting:
            failed = [s.reconciliation for s in sheet_results if not s.all_accounting_passed]
            if failed:
                raise RowAccountingError(failed)

        return wb_result


def inspect_workbook(
    source: ExcelSource,
    *,
    sheet_names: Sequence[str] | None = None,
    config: ParserConfig | None = None,
    **config_overrides,
) -> WorkbookInspection:
    """Stateless convenience function to inspect an `.xlsx` workbook (< 1s, no LLM calls)."""
    parser = ExcelParser(config=config, **config_overrides)
    return parser.inspect(source, sheet_names=sheet_names)


def extract_workbook(
    source: ExcelSource,
    *,
    sheet_names: Sequence[str] | None = None,
    fail_on_accounting: bool = False,
    output_dir: str | Path | None = None,
    pg_table: str = "equipment_catalog",
    save_debug_excel: bool = False,
    on_progress: Callable[[ProgressEvent], None] | None = None,
    config: ParserConfig | None = None,
    client: VLLMClient | None = None,
    **config_overrides,
) -> WorkbookExtractionResult:
    """Stateless convenience function to extract & reconcile an `.xlsx` workbook."""
    parser = ExcelParser(config=config, client=client, **config_overrides)
    return parser.extract(
        source,
        sheet_names=sheet_names,
        fail_on_accounting=fail_on_accounting,
        output_dir=output_dir,
        pg_table=pg_table,
        save_debug_excel=save_debug_excel,
        on_progress=on_progress,
    )
