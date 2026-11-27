# Step 4: Error and value normalization

import math
import re
import unicodedata
from dataclasses import dataclass, field, replace
from datetime import date, datetime, time
from typing import Any, Literal

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from excel_parsing.preprocess.decompose import CellNode, SheetGrid
from excel_parsing.preprocess.detection import Region

ErrorStrategy = Literal["tag", "null", "keep"]

EXCEL_ERROR_MEANINGS: dict[str, str] = {
    "#DIV/0!": "Division by zero",
    "#N/A": "Value not available",
    "#NAME?": "Unrecognized formula name",
    "#NULL!": "Empty range intersection",
    "#NUM!": "Invalid numeric value",
    "#REF!": "Invalid cell reference",
    "#VALUE!": "Wrong value or operand type",
    "#SPILL!": "Dynamic array spill blocked",
    "#CALC!": "Calculation engine error",
    "#GETTING_DATA": "External data still loading",
    "#FIELD!": "Invalid linked data field",
    "#CONNECT!": "External connection error",
    "#BLOCKED!": "Blocked external data source",
    "#UNKNOWN!": "Unknown function or error",
}

_EXCEL_CR_RE = re.compile(r"_x000D_\r?\n?|\r\n|\r")
_MULTI_SPACE_RE = re.compile(r"[ \t\u3000]+")


@dataclass(frozen=True)
class CellError:
    sheet_name: str
    row: int
    col: int
    error_code: str
    meaning: str
    formula: str | None = None
    is_merge_origin: bool = True

    @property
    def coord(self) -> str:
        return f"{get_column_letter(self.col)}{self.row}"


@dataclass
class NormalizationReport:
    sheet_name: str
    errors: list[CellError] = field(default_factory=list)
    blank_strings_removed: int = 0
    strings_cleaned: int = 0
    non_finite_floats: int = 0

    @property
    def error_count(self) -> int:
        return len(self.errors)

    def error_summary(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for err in self.errors:
            counts[err.error_code] = counts.get(err.error_code, 0) + 1
        return counts


def detect_excel_error(value: Any) -> tuple[str, str] | None:
    """Return `(error_code, meaning)` if `value` is an Excel error or non-finite float."""
    if isinstance(value, str):
        code = value.strip().upper()
        if code in EXCEL_ERROR_MEANINGS:
            return (code, EXCEL_ERROR_MEANINGS[code])
    elif isinstance(value, float):
        if math.isnan(value):
            return ("#NUM!", "NaN floating-point value")
        if math.isinf(value):
            return ("#NUM!", "Infinite floating-point value")
    return None


def _clean_string(
    text: str,
    *,
    normalize_unicode: bool = False,
    collapse_spaces: bool = False,
) -> str:
    cleaned = _EXCEL_CR_RE.sub("\n", text)
    if normalize_unicode:
        cleaned = unicodedata.normalize("NFKC", cleaned)
    if collapse_spaces:
        cleaned = "\n".join(
            _MULTI_SPACE_RE.sub(" ", line).strip() for line in cleaned.split("\n")
        )
    else:
        cleaned = cleaned.strip()
    return cleaned


def _format_temporal(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.isoformat(timespec="seconds")
    return value


def _lookup_formula(formula_ws: Worksheet | None, node: CellNode) -> str | None:
    if formula_ws is None:
        return None
    r, c = node.merge.origin if node.merge is not None else (node.row, node.col)
    raw = formula_ws.cell(r, c).value
    if isinstance(raw, str) and raw.startswith("="):
        return raw
    return None


def normalize_grid(
    grid: SheetGrid,
    *,
    error_strategy: ErrorStrategy = "tag",
    drop_blank_strings: bool = True,
    normalize_unicode: bool = False,
    collapse_spaces: bool = False,
    format_dates_iso: bool = False,
    formula_ws: Worksheet | None = None,
) -> tuple[SheetGrid, NormalizationReport]:
    """Return a normalized copy of `grid` and a `NormalizationReport`.

    Args:
        grid: Resolved SheetGrid from Step 1.
        error_strategy: How to handle Excel error values (`#DIV/0!`, `#REF!`, etc.):
            - `"tag"` (default): replace with `"[ERROR:#DIV/0!]"` so LLMs see the error explicitly
            - `"null"`: replace with `None` (drops cell from sparse grid)
            - `"keep"`: keep the original error string as-is
        drop_blank_strings: If True, whitespace-only strings are treated as empty and dropped.
        normalize_unicode: If True, apply Unicode NFKC normalization (e.g. full-width -> half-width).
        collapse_spaces: If True, collapse runs of spaces/tabs within each line.
        format_dates_iso: If True, convert `date`/`datetime`/`time` values to ISO-8601 strings.
        formula_ws: Optional worksheet loaded with `data_only=False` to attach the source
            formula to any recorded `CellError`.
    """
    report = NormalizationReport(sheet_name=grid.sheet_name)
    cleaned_cells: dict[tuple[int, int], CellNode] = {}

    for coord, node in grid.cells.items():
        val = node.value

        err = detect_excel_error(val)
        if err is not None:
            code, meaning = err
            if isinstance(val, float):
                report.non_finite_floats += 1
            report.errors.append(
                CellError(
                    sheet_name=grid.sheet_name,
                    row=node.row,
                    col=node.col,
                    error_code=code,
                    meaning=meaning,
                    formula=_lookup_formula(formula_ws, node),
                    is_merge_origin=node.is_merge_origin,
                )
            )
            if error_strategy == "null":
                continue
            new_val = f"[ERROR:{code}]" if error_strategy == "tag" else val
            cleaned_cells[coord] = replace(node, value=new_val)
            continue

        if isinstance(val, str):
            new_str = _clean_string(
                val,
                normalize_unicode=normalize_unicode,
                collapse_spaces=collapse_spaces,
            )
            if not new_str and drop_blank_strings:
                report.blank_strings_removed += 1
                continue
            if new_str != val:
                report.strings_cleaned += 1
            cleaned_cells[coord] = replace(node, value=new_str)
            continue

        if format_dates_iso and isinstance(val, (datetime, date, time)):
            cleaned_cells[coord] = replace(node, value=_format_temporal(val))
            continue

        cleaned_cells[coord] = node

    normalized_grid = SheetGrid(
        sheet_name=grid.sheet_name,
        max_row=grid.max_row,
        max_col=grid.max_col,
        cells=cleaned_cells,
        merges=list(grid.merges),
        hidden_rows=set(grid.hidden_rows),
        hidden_cols=set(grid.hidden_cols),
    )
    return normalized_grid, report


def normalize_regions(
    regions: list[Region],
    normalized_grid: SheetGrid,
) -> list[Region]:
    """Update a list of detected `Region`s to use cells from `normalized_grid`.

    Regions that become completely empty after dropping blank/error cells are omitted.
    """
    updated: list[Region] = []
    for region in regions:
        new_cells = {
            coord: normalized_grid.cells[coord]
            for coord in region.cells
            if coord in normalized_grid.cells
        }
        if not new_cells:
            continue
        rows = [r for r, _ in new_cells]
        cols = [c for _, c in new_cells]
        updated.append(
            Region(
                region_id=region.region_id,
                sheet_name=region.sheet_name,
                min_row=min(rows),
                min_col=min(cols),
                max_row=max(rows),
                max_col=max(cols),
                cells=new_cells,
            )
        )
    return updated
