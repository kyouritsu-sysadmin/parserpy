"""`excel_parsing` — Layer 1 Python SDK for deterministic Excel decomposition,
stateful vLLM extraction, row-accounting reconciliation, and PostgreSQL export.
"""

from excel_parsing.exceptions import (
    ExcelParsingError,
    InvalidWorkbookError,
    RowAccountingError,
    SheetNotFoundError,
    VLLMHTTPError,
    VLLMOutputTruncatedError,
)
from excel_parsing.exporters.postgres import (
    PostgresRecordsetPayload,
    build_pg_recordset_payload,
    denormalize_extraction,
)
from excel_parsing.extraction.extract import VLLMClient
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
from excel_parsing.pipeline import (
    ExcelParser,
    extract_workbook,
    inspect_workbook,
    is_extractable_region,
    preprocess_sheet,
    preprocess_workbook,
)
from excel_parsing.reconcile.reconcile import (
    IgnoredRow,
    ImportCategory,
    ImportExtraction,
    ImportItem,
    RegionReconciliation,
    RowAccountingResult,
    RowValidationIssue,
    SheetReconciliation,
)

__version__ = "0.2.0"

__all__ = [
    "ExcelParser",
    "ExcelParsingError",
    "ExcelSource",
    "IgnoredRow",
    "ImportCategory",
    "ImportExtraction",
    "ImportItem",
    "InvalidWorkbookError",
    "ParserConfig",
    "PostgresRecordsetPayload",
    "PreprocessedSheet",
    "ProgressEvent",
    "RegionReconciliation",
    "RegionSummary",
    "RowAccountingError",
    "RowAccountingResult",
    "RowValidationIssue",
    "SheetExtractionResult",
    "SheetInspection",
    "SheetNotFoundError",
    "SheetReconciliation",
    "VLLMClient",
    "VLLMHTTPError",
    "VLLMOutputTruncatedError",
    "WorkbookExtractionResult",
    "WorkbookInspection",
    "build_pg_recordset_payload",
    "denormalize_extraction",
    "extract_workbook",
    "inspect_workbook",
    "is_extractable_region",
    "preprocess_sheet",
    "preprocess_workbook",
]
