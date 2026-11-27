# Step 3: Formatting feature extraction

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.worksheet.worksheet import Worksheet

from excel_parsing.preprocess.decompose import CellNode, SheetGrid
from excel_parsing.preprocess.detection import Region

FormatType = Literal["percent", "currency", "date", "number", "text", "general"]
RowRole = Literal["header", "total", "section", "data"]

_EMPHASIS_BORDERS = frozenset({"medium", "thick", "double"})
_CURRENCY_RE = re.compile(r"[$€£¥￥₩]|USD|JPY|EUR|GBP")
_DATE_TOKEN_RE = re.compile(r"[ymdhs]", re.IGNORECASE)
_BRACKET_OR_QUOTE_RE = re.compile(r'\[[^\]]*\]|"[^"]*"|\\.')
_TOTAL_KEYWORDS = (
    "合計",
    "小計",
    "総計",
    "材比",
    "掛率",
    "値引",
    "消費税",
    "諸経費",
    "total",
    "subtotal",
    "grand total",
)


@dataclass(frozen=True)
class BorderInfo:
    top: str | None = None
    bottom: str | None = None
    left: str | None = None
    right: str | None = None

    @property
    def has_any(self) -> bool:
        return any((self.top, self.bottom, self.left, self.right))

    @property
    def has_bottom_emphasis(self) -> bool:
        return self.bottom in _EMPHASIS_BORDERS

    @property
    def has_top_emphasis(self) -> bool:
        return self.top in _EMPHASIS_BORDERS


@dataclass(frozen=True)
class CellFormat:
    is_bold: bool = False
    is_italic: bool = False
    font_size: float | None = None
    font_color: str | None = None
    fill_color: str | None = None
    borders: BorderInfo = BorderInfo()
    horizontal: str | None = None
    vertical: str | None = None
    indent: int = 0
    wrap_text: bool = False
    number_format: str = "General"
    format_type: FormatType = "general"

    @property
    def has_fill(self) -> bool:
        return self.fill_color is not None

    def to_compact_dict(self) -> dict[str, Any]:
        """Return only non-default formatting signals for compact inspection/serialization."""
        out: dict[str, Any] = {}
        if self.is_bold:
            out["bold"] = True
        if self.is_italic:
            out["italic"] = True
        if self.font_size is not None:
            out["font_size"] = self.font_size
        if self.font_color is not None:
            out["font_color"] = self.font_color
        if self.fill_color is not None:
            out["fill_color"] = self.fill_color
        if self.borders.has_any:
            out["borders"] = {
                k: v
                for k, v in (
                    ("top", self.borders.top),
                    ("bottom", self.borders.bottom),
                    ("left", self.borders.left),
                    ("right", self.borders.right),
                )
                if v is not None
            }
        if self.horizontal is not None:
            out["align"] = self.horizontal
        if self.indent > 0:
            out["indent"] = self.indent
        if self.format_type != "general":
            out["format_type"] = self.format_type
        return out


def _normalize_color(color: Color | None) -> str | None:
    if color is None:
        return None

    if color.type == "rgb":
        rgb = color.rgb
        if not rgb or rgb in ("00000000", "000000"):
            return None
        if len(rgb) == 8:
            # ARGB -> #RRGGBB (ignore fully transparent 00 alpha)
            if rgb[:2] == "00":
                return None
            return f"#{rgb[2:].upper()}"
        return f"#{rgb.upper()}"

    if color.type == "theme":
        tint = getattr(color, "tint", 0.0) or 0.0
        if abs(tint) > 1e-6:
            return f"theme:{color.theme}({tint:+.2f})"
        return f"theme:{color.theme}"

    if color.type == "indexed":
        # 64 and 65 are Excel's automatic/system foreground and background colors
        if color.indexed in (None, 64, 65):
            return None
        return f"indexed:{color.indexed}"

    return None


def _classify_number_format(number_format: str | None) -> FormatType:
    if not number_format or number_format.lower() == "general":
        return "general"
    if number_format == "@":
        return "text"

    # Inspect only the positive (first) section of multi-part Excel formats
    primary = number_format.split(";")[0]
    cleaned = _BRACKET_OR_QUOTE_RE.sub("", primary)

    if "%" in cleaned:
        return "percent"
    if _CURRENCY_RE.search(primary):
        return "currency"
    if _DATE_TOKEN_RE.search(cleaned):
        return "date"
    if "0" in cleaned or "#" in cleaned:
        return "number"
    return "general"


def _extract_cell_format(ws: Worksheet, node: CellNode) -> CellFormat:
    """Extract formatting for `node`, handling merged-cell origin vs perimeter borders."""
    if node.merge is not None:
        span = node.merge
        style_cell = ws.cell(span.min_row, span.min_col)
        top_border = ws.cell(span.min_row, node.col).border.top
        bottom_border = ws.cell(span.max_row, node.col).border.bottom
        left_border = ws.cell(node.row, span.min_col).border.left
        right_border = ws.cell(node.row, span.max_col).border.right
    else:
        style_cell = ws.cell(node.row, node.col)
        top_border = style_cell.border.top
        bottom_border = style_cell.border.bottom
        left_border = style_cell.border.left
        right_border = style_cell.border.right

    font = style_cell.font
    fill = style_cell.fill
    alignment = style_cell.alignment

    fill_color: str | None = None
    if fill is not None and fill.fill_type == "solid":
        fill_color = _normalize_color(fill.fgColor) or _normalize_color(fill.start_color)

    borders = BorderInfo(
        top=top_border.style if top_border is not None else None,
        bottom=bottom_border.style if bottom_border is not None else None,
        left=left_border.style if left_border is not None else None,
        right=right_border.style if right_border is not None else None,
    )

    num_fmt = style_cell.number_format or "General"

    return CellFormat(
        is_bold=bool(font.bold) if font is not None else False,
        is_italic=bool(font.italic) if font is not None else False,
        font_size=float(font.size) if font is not None and font.size is not None else None,
        font_color=_normalize_color(font.color) if font is not None else None,
        fill_color=fill_color,
        borders=borders,
        horizontal=alignment.horizontal if alignment is not None else None,
        vertical=alignment.vertical if alignment is not None else None,
        indent=int(alignment.indent or 0) if alignment is not None else 0,
        wrap_text=bool(alignment.wrap_text) if alignment is not None else False,
        number_format=num_fmt,
        format_type=_classify_number_format(num_fmt),
    )


def extract_formatting(
    ws: Worksheet,
    grid: SheetGrid,
) -> dict[tuple[int, int], CellFormat]:
    """Extract normalized visual formatting features for every cell in `grid`."""
    return {
        coord: _extract_cell_format(ws, node)
        for coord, node in grid.cells.items()
    }


def classify_region_rows(
    region: Region,
    formats: dict[tuple[int, int], CellFormat],
) -> dict[int, RowRole]:
    """Infer lightweight row roles (`section`, `header`, `total`, `data`) for a Region.

    Uses visual cues (bold, fill, borders, merge spans) plus simple structural
    heuristics so downstream serialization can highlight headers and totals.
    """
    roles: dict[int, RowRole] = {}

    for r in range(region.min_row, region.max_row + 1):
        row_nodes = [
            region.cells[(r, c)]
            for c in range(region.min_col, region.max_col + 1)
            if (r, c) in region.cells
        ]
        if not row_nodes:
            continue

        row_fmts = [formats.get((n.row, n.col), CellFormat()) for n in row_nodes]
        count = len(row_nodes)

        # 1. Section banner: single wide horizontal merge covering most of the region
        unique_merges = {n.merge for n in row_nodes if n.merge is not None}
        if (
            len(unique_merges) == 1
            and region.num_cols >= 3
            and next(iter(unique_merges)).col_span >= max(2, int(region.num_cols * 0.7))
        ):
            roles[r] = "section"
            continue

        bold_ratio = sum(1 for f in row_fmts if f.is_bold) / count
        fill_ratio = sum(1 for f in row_fmts if f.has_fill) / count
        bottom_emph_ratio = sum(1 for f in row_fmts if f.borders.has_bottom_emphasis) / count
        top_emph_ratio = sum(1 for f in row_fmts if f.borders.has_top_emphasis) / count
        string_ratio = sum(1 for n in row_nodes if isinstance(n.value, str)) / count

        has_total_kw = any(
            isinstance(n.value, str)
            and any(
                kw in unicodedata.normalize("NFKC", n.value).strip().lower()
                for kw in _TOTAL_KEYWORDS
            )
            for n in row_nodes
        )

        # Single-cell or bracketed section header row (e.g. "【DS】", "【高圧機器】", "ACB(ZCT付) 40")
        origin_nodes = [n for n in row_nodes if n.is_merge_origin] or row_nodes
        if len(origin_nodes) == 1 and isinstance(origin_nodes[0].value, str):
            norm_txt = unicodedata.normalize("NFKC", origin_nodes[0].value).strip()
            if (norm_txt.startswith("【") and "】" in norm_txt) or (
                region.num_cols >= 2 and r > region.min_row and (row_fmts[0].is_bold or row_fmts[0].has_fill)
            ):
                roles[r] = "section"
                continue

        # 2. Total / subtotal row
        if r > region.min_row and (
            bottom_emph_ratio >= 0.5
            or (has_total_kw and (bold_ratio >= 0.3 or fill_ratio >= 0.3 or top_emph_ratio >= 0.3))
        ):
            roles[r] = "total"
            continue

        # 3. Header row (top of region or contiguous with previous header/section)
        prev_role = roles.get(r - 1)
        is_top_band = r == region.min_row or prev_role in ("header", "section")
        if is_top_band and string_ratio >= 0.6 and (bold_ratio >= 0.5 or fill_ratio >= 0.5):
            roles[r] = "header"
            continue

        roles[r] = "data"

    return roles


_ROLE_FILLS: dict[RowRole, PatternFill] = {
    "section": PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid"),
    "header": PatternFill(start_color="CFE2F3", end_color="CFE2F3", fill_type="solid"),
    "total": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "data": PatternFill(fill_type=None),
}


def _apply_cell_format(cell: Any, fmt: CellFormat, *, apply_fill: bool = True) -> None:
    font_kwargs: dict[str, Any] = {
        "bold": fmt.is_bold,
        "italic": fmt.is_italic,
    }
    if fmt.font_size is not None:
        font_kwargs["size"] = fmt.font_size
    if fmt.font_color is not None and fmt.font_color.startswith("#"):
        font_kwargs["color"] = fmt.font_color[1:]
    cell.font = Font(**font_kwargs)

    if apply_fill and fmt.fill_color is not None and fmt.fill_color.startswith("#"):
        hex_color = fmt.fill_color[1:]
        cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    if fmt.borders.has_any:
        cell.border = Border(
            top=Side(style=fmt.borders.top) if fmt.borders.top else None,
            bottom=Side(style=fmt.borders.bottom) if fmt.borders.bottom else None,
            left=Side(style=fmt.borders.left) if fmt.borders.left else None,
            right=Side(style=fmt.borders.right) if fmt.borders.right else None,
        )

    cell.alignment = Alignment(
        horizontal=fmt.horizontal,
        vertical=fmt.vertical,
        indent=fmt.indent,
        wrap_text=fmt.wrap_text,
    )
    if fmt.number_format:
        cell.number_format = fmt.number_format


def save_formatted_regions_to_excel(
    regions: list[Region],
    formats: dict[tuple[int, int], CellFormat],
    output_path: str | Path,
    *,
    include_overlay: bool = True,
    include_region_sheets: bool = True,
) -> Path:
    """Save formatting extraction and row-role results into a single Excel file.

    Sheets created:
    - `_summary`: per-region counts of inferred row roles and extracted visual signals
    - `_roles_overlay` (if `include_overlay=True`): original `(row, col)` layout
      color-coded by row role (`section`=purple, `header`=blue, `total`=yellow)
      with extracted bold/italic, borders, alignment, and number formats applied
    - `R{id}_{range}` (if `include_region_sheets=True`): each region cropped to A1
      with column A showing its inferred `_role` and cells styled using extracted formats
    """
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    summary_ws = wb.active
    summary_ws.title = "_summary"

    headers = [
        "region_id",
        "coord_range",
        "kind",
        "section_rows",
        "header_rows",
        "total_rows",
        "data_rows",
        "bold_cells",
        "filled_cells",
        "formatted_cells",
    ]
    summary_ws.append(headers)
    bold_font = Font(bold=True)
    for col_idx in range(1, len(headers) + 1):
        summary_ws.cell(1, col_idx).font = bold_font

    region_roles: dict[int, dict[int, RowRole]] = {}

    for region in regions:
        roles = classify_region_rows(region, formats)
        region_roles[region.region_id] = roles

        role_counts: dict[RowRole, int] = {"section": 0, "header": 0, "total": 0, "data": 0}
        for role in roles.values():
            role_counts[role] += 1

        reg_fmts = [formats.get(coord, CellFormat()) for coord in region.cells]
        summary_ws.append(
            [
                region.region_id,
                region.coord_range,
                region.kind,
                role_counts["section"],
                role_counts["header"],
                role_counts["total"],
                role_counts["data"],
                sum(1 for f in reg_fmts if f.is_bold),
                sum(1 for f in reg_fmts if f.has_fill),
                sum(1 for f in reg_fmts if f.format_type != "general"),
            ]
        )

    if include_overlay:
        overlay_ws = wb.create_sheet("_roles_overlay")
        for region in regions:
            roles = region_roles[region.region_id]
            for (r_idx, c_idx), node in region.cells.items():
                cell = overlay_ws.cell(r_idx, c_idx)
                cell.value = node.value
                fmt = formats.get((r_idx, c_idx), CellFormat())
                _apply_cell_format(cell, fmt, apply_fill=False)
                role = roles.get(r_idx, "data")
                if role != "data":
                    cell.fill = _ROLE_FILLS[role]

    if include_region_sheets:
        for region in regions:
            roles = region_roles[region.region_id]
            safe_range = region.coord_range.replace(":", "_")
            title = f"R{region.region_id:02d}_{safe_range}"[:31]
            ws = wb.create_sheet(title)

            ws.cell(1, 1, "_role").font = bold_font
            for c_offset in range(region.num_cols):
                ws.cell(1, c_offset + 2, f"col_{c_offset + 1}").font = bold_font

            for r_offset, r_idx in enumerate(range(region.min_row, region.max_row + 1), start=2):
                role = roles.get(r_idx, "data")
                role_cell = ws.cell(r_offset, 1, role)
                role_cell.font = bold_font
                if role != "data":
                    role_cell.fill = _ROLE_FILLS[role]

                for c_offset, c_idx in enumerate(range(region.min_col, region.max_col + 1), start=2):
                    node = region.cells.get((r_idx, c_idx))
                    if node is None:
                        continue
                    cell = ws.cell(r_offset, c_offset, node.value)
                    fmt = formats.get((r_idx, c_idx), CellFormat())
                    _apply_cell_format(cell, fmt, apply_fill=True)

    wb.save(out_path)
    return out_path

