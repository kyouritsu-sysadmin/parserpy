"""Stage 1: Spatial decomposition, normalization, region detection, and formatting."""

from excel_parsing.preprocess.decompose import (
    CellNode,
    MergeSpan,
    SheetGrid,
    resolve_merged_cells,
    unmerge_worksheet_inplace,
)
from excel_parsing.preprocess.detection import (
    Region,
    detect_regions,
    save_regions_to_excel,
)
from excel_parsing.preprocess.formatting import (
    BorderInfo,
    CellFormat,
    FormatType,
    RowRole,
    classify_region_rows,
    extract_formatting,
    save_formatted_regions_to_excel,
)
from excel_parsing.preprocess.normalize import (
    CellError,
    ErrorStrategy,
    NormalizationReport,
    detect_excel_error,
    normalize_grid,
    normalize_regions,
)

__all__ = [
    "BorderInfo",
    "CellError",
    "CellFormat",
    "CellNode",
    "ErrorStrategy",
    "FormatType",
    "MergeSpan",
    "NormalizationReport",
    "Region",
    "RowRole",
    "SheetGrid",
    "classify_region_rows",
    "detect_excel_error",
    "detect_regions",
    "extract_formatting",
    "normalize_grid",
    "normalize_regions",
    "resolve_merged_cells",
    "save_formatted_regions_to_excel",
    "save_regions_to_excel",
    "unmerge_worksheet_inplace",
]
