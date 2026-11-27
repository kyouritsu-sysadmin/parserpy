"""SDK configuration, inspection, progress, and extraction result models."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, BinaryIO, Literal

from excel_parsing.exporters.postgres import (
    JsonFuncType,
    ParamStyle,
    PostgresRecordsetPayload,
    QueryMode,
    build_pg_recordset_payload,
)
from excel_parsing.extraction.extract import ExtractionResult
from excel_parsing.preprocess.decompose import SheetGrid
from excel_parsing.preprocess.detection import Region
from excel_parsing.preprocess.formatting import CellFormat, RowRole
from excel_parsing.preprocess.normalize import ErrorStrategy, NormalizationReport
from excel_parsing.reconcile.reconcile import ImportExtraction, RegionReconciliation, SheetReconciliation

ExcelSource = str | Path | bytes | bytearray | BinaryIO


@dataclass(frozen=True)
class ParserConfig:
    """Configuration for preprocessing, vLLM extraction, and reconciliation."""

    # vLLM client settings
    vllm_base_url: str = "http://localhost:8005/v1"
    vllm_model: str | None = None
    vllm_api_key: str = "EMPTY"
    vllm_timeout: float = 180.0
    default_max_tokens: int = 8192
    disable_thinking: bool = True
    max_retries: int = 2
    max_concurrent_requests: int = 6

    # Preprocessing & decomposition settings
    load_formulas: bool = False
    skip_hidden: bool = True
    connectivity: Literal[4, 8] = 4
    row_gap: int = 1
    col_gap: int = 0
    merge_overlapping_bboxes: bool = False
    split_banner_bridges: bool = True
    split_side_by_side: bool = True
    reunite_split_columns: bool = True
    min_cells: int = 1
    error_strategy: ErrorStrategy = "tag"
    normalize_unicode: bool = True
    collapse_spaces: bool = True
    format_dates_iso: bool = True

    # Chunking & concurrency settings
    max_rows_per_chunk: int = 25
    max_chunk_workers: int = 3
    max_region_workers: int = 6
    include_coords: bool = True
    retry_missing: bool = True


@dataclass
class PreprocessedSheet:
    """In-memory Stage 1 preprocessing output for a single worksheet."""

    sheet_name: str
    grid: SheetGrid
    regions: list[Region]
    formats: dict[tuple[int, int], CellFormat]
    row_roles: dict[int, dict[int, RowRole]]
    report: NormalizationReport


@dataclass(frozen=True)
class RegionSummary:
    """Lightweight metadata for a detected region in Stages 1-2."""

    region_id: int
    coord_range: str
    kind: str
    rows: int
    cols: int
    cell_count: int
    density: float
    extractable: bool


@dataclass
class SheetInspection:
    """Fast CPU-only Stage 1-2 inspection output for one worksheet."""

    sheet_name: str
    max_row: int
    max_col: int
    visible_cells: int
    merges: int
    hidden_rows: int
    hidden_cols: int
    errors: dict[str, int]
    blank_strings_removed: int
    regions: list[RegionSummary]
    compact_html: str

    @property
    def extractable_region_count(self) -> int:
        return sum(1 for r in self.regions if r.extractable)

    def to_dict(self, *, include_html: bool = False) -> dict[str, Any]:
        out: dict[str, Any] = {
            "sheet_name": self.sheet_name,
            "max_row": self.max_row,
            "max_col": self.max_col,
            "visible_cells": self.visible_cells,
            "merges": self.merges,
            "hidden_rows": self.hidden_rows,
            "hidden_cols": self.hidden_cols,
            "errors": dict(self.errors),
            "blank_strings_removed": self.blank_strings_removed,
            "total_regions": len(self.regions),
            "extractable_regions": self.extractable_region_count,
            "compact_html_chars": len(self.compact_html),
            "regions": [
                {
                    "region_id": r.region_id,
                    "coord_range": r.coord_range,
                    "kind": r.kind,
                    "rows": r.rows,
                    "cols": r.cols,
                    "cell_count": r.cell_count,
                    "density": round(r.density, 4),
                    "extractable": r.extractable,
                }
                for r in self.regions
            ],
        }
        if include_html:
            out["compact_html"] = self.compact_html
        return out


@dataclass
class WorkbookInspection:
    """Fast CPU-only Stage 1-2 inspection output for an entire workbook."""

    all_sheet_names: list[str]
    visible_sheet_names: list[str]
    sheets: list[SheetInspection]
    elapsed_seconds: float

    def to_dict(self, *, include_html: bool = False) -> dict[str, Any]:
        return {
            "all_sheet_names": list(self.all_sheet_names),
            "visible_sheet_names": list(self.visible_sheet_names),
            "elapsed_seconds": self.elapsed_seconds,
            "sheets": [s.to_dict(include_html=include_html) for s in self.sheets],
        }


@dataclass(frozen=True)
class ProgressEvent:
    """Structured progress callback event emitted during extraction."""

    sheet_name: str
    completed_regions: int
    total_regions: int
    region_id: int
    coord_range: str
    status: Literal["extracted", "failed"]
    valid_items: int = 0
    validation_issues: int = 0
    accounting_passed: bool = False
    elapsed_seconds: float = 0.0
    error: str | None = None


@dataclass
class SheetExtractionResult:
    """Full Stage 1-4 output for a single worksheet."""

    sheet_name: str
    preprocessed: PreprocessedSheet
    compact_html: str
    region_extractions: list[ExtractionResult]
    reconciliation: SheetReconciliation
    elapsed_seconds: float
    saved_files: dict[str, Path] = field(default_factory=dict)

    @property
    def extraction(self) -> ImportExtraction | None:
        return self.reconciliation.extraction

    @property
    def all_accounting_passed(self) -> bool:
        return self.reconciliation.all_accounting_passed

    @property
    def region_reconciliations(self) -> list[RegionReconciliation]:
        return self.reconciliation.regions

    def to_pg_recordset(
        self,
        *,
        table: str = "equipment_catalog",
        mode: QueryMode = "insert",
        include_sheet_name: bool = False,
        include_fields: Sequence[str] | None = None,
        column_map: Mapping[str, str] | None = None,
        extra_columns: Mapping[str, str] | None = None,
        param_style: ParamStyle = "dollar",
        json_type: JsonFuncType = "json",
        on_conflict: str | None = None,
        returning: str | None = None,
        alias: str = "r",
    ) -> PostgresRecordsetPayload:
        return build_pg_recordset_payload(
            self.reconciliation,
            table=table,
            mode=mode,
            include_sheet_name=include_sheet_name,
            include_fields=include_fields,
            column_map=column_map,
            extra_columns=extra_columns,
            param_style=param_style,
            json_type=json_type,
            on_conflict=on_conflict,
            returning=returning,
            alias=alias,
        )


@dataclass
class WorkbookExtractionResult:
    """Full Stage 1-4 output for a workbook across all processed sheets."""

    sheets: list[SheetExtractionResult]
    content_sha256: str
    elapsed_seconds: float

    @property
    def all_accounting_passed(self) -> bool:
        return all(s.all_accounting_passed for s in self.sheets)

    @property
    def total_valid_items(self) -> int:
        return sum(
            len(s.extraction.items) for s in self.sheets if s.extraction is not None
        )

    def get_sheet(self, sheet_name: str) -> SheetExtractionResult:
        for s in self.sheets:
            if s.sheet_name == sheet_name:
                return s
        raise KeyError(f"Sheet {sheet_name!r} not found in extraction result")

    def to_pg_recordset(
        self,
        *,
        table: str = "equipment_catalog",
        mode: QueryMode = "insert",
        include_sheet_name: bool | None = None,
        include_fields: Sequence[str] | None = None,
        column_map: Mapping[str, str] | None = None,
        extra_columns: Mapping[str, str] | None = None,
        param_style: ParamStyle = "dollar",
        json_type: JsonFuncType = "json",
        on_conflict: str | None = None,
        returning: str | None = None,
        alias: str = "r",
    ) -> PostgresRecordsetPayload:
        return build_pg_recordset_payload(
            [s.reconciliation for s in self.sheets],
            table=table,
            mode=mode,
            include_sheet_name=include_sheet_name,
            include_fields=include_fields,
            column_map=column_map,
            extra_columns=extra_columns,
            param_style=param_style,
            json_type=json_type,
            on_conflict=on_conflict,
            returning=returning,
            alias=alias,
        )
