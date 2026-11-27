"""Typed SDK exceptions for `excel_parsing`."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from excel_parsing.reconcile.reconcile import SheetReconciliation


class ExcelParsingError(Exception):
    """Base exception for all `excel_parsing` SDK errors."""


class InvalidWorkbookError(ExcelParsingError, ValueError):
    """Raised when the input bytes, stream, or path is not a valid `.xlsx` workbook."""


class SheetNotFoundError(ExcelParsingError, KeyError):
    """Raised when a requested worksheet name does not exist in the workbook."""


class VLLMOutputTruncatedError(ExcelParsingError, RuntimeError):
    """Raised when vLLM hits `max_tokens` and returns empty or truncated JSON."""


class VLLMHTTPError(ExcelParsingError, RuntimeError):
    """Raised when vLLM returns a non-2xx HTTP status."""

    def __init__(self, code: int, body: str) -> None:
        super().__init__(f"vLLM HTTP {code}: {body[:400]}")
        self.code = code
        self.body = body


class RowAccountingError(ExcelParsingError, RuntimeError):
    """Raised when `fail_on_accounting=True` and one or more sheets fail row accounting."""

    def __init__(self, failed_sheets: list[SheetReconciliation]) -> None:
        summary = ", ".join(
            f"{s.sheet_name} (issues={sum(len(r.validation_issues) for r in s.regions)})"
            for s in failed_sheets
        )
        super().__init__(f"Row-accounting invariant failed on sheet(s): {summary}")
        self.failed_sheets = failed_sheets
