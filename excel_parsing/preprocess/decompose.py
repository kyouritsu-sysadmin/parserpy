# Spatial decomposition of excel files based on the cells

from dataclasses import dataclass, field
from typing import Any, Literal

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


# Step 1: Merge cell resolution and propagation


@dataclass(frozen=True)
class MergeSpan:
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def row_span(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def col_span(self) -> int:
        return self.max_col - self.min_col + 1

    @property
    def origin(self) -> tuple[int, int]:
        return (self.min_row, self.min_col)

    @property
    def orientation(self) -> Literal["horizontal", "vertical", "block", "single"]:
        if self.row_span == 1 and self.col_span > 1:
            return "horizontal"
        if self.col_span == 1 and self.row_span > 1:
            return "vertical"
        if self.row_span > 1 and self.col_span > 1:
            return "block"
        return "single"

    @property
    def coord_range(self) -> str:
        start = f"{get_column_letter(self.min_col)}{self.min_row}"
        end = f"{get_column_letter(self.max_col)}{self.max_row}"
        return f"{start}:{end}"


@dataclass
class CellNode:
    row: int
    col: int
    raw_value: Any
    value: Any
    merge: MergeSpan | None = None
    is_merge_origin: bool = True

    @property
    def coord(self) -> str:
        return f"{get_column_letter(self.col)}{self.row}"


@dataclass
class SheetGrid:
    sheet_name: str
    max_row: int
    max_col: int
    cells: dict[tuple[int, int], CellNode] = field(default_factory=dict)
    merges: list[MergeSpan] = field(default_factory=list)
    hidden_rows: set[int] = field(default_factory=set)
    hidden_cols: set[int] = field(default_factory=set)

    def get(self, row: int, col: int) -> CellNode | None:
        return self.cells.get((row, col))


def _collect_hidden_axes(ws: Worksheet) -> tuple[set[int], set[int]]:
    hidden_rows: set[int] = {
        int(r)
        for r, dim in ws.row_dimensions.items()
        if r is not None and bool(getattr(dim, "hidden", False))
    }
    hidden_cols: set[int] = set()
    for key, dim in ws.column_dimensions.items():
        if not bool(getattr(dim, "hidden", False)):
            continue
        min_c = getattr(dim, "min", None)
        max_c = getattr(dim, "max", None)
        if isinstance(min_c, int) and isinstance(max_c, int):
            hidden_cols.update(range(min_c, max_c + 1))
        elif isinstance(key, str):
            try:
                hidden_cols.add(column_index_from_string(key))
            except ValueError:
                pass
    return hidden_rows, hidden_cols


def resolve_merged_cells(
    ws: Worksheet,
    *,
    propagate_empty: bool = False,
    skip_hidden: bool = True,
) -> SheetGrid:
    """Resolve all merged cell ranges into a non-destructive SheetGrid.

    Each cell inside a merged range receives:
    - `value`: propagated top-left origin value
    - `raw_value`: original cell value (None for non-origin merged cells)
    - `merge`: MergeSpan metadata describing the visible merged block
    - `is_merge_origin`: True only for the top-left visible cell of the merge

    Hidden rows and hidden `ColumnDimension` ranges (`min..max`) are excluded when
    `skip_hidden=True` so hidden template columns do not bloat downstream LLM prompts.
    """
    hidden_rows, hidden_cols = _collect_hidden_axes(ws) if skip_hidden else (set(), set())

    grid = SheetGrid(
        sheet_name=ws.title,
        max_row=ws.max_row or 0,
        max_col=ws.max_column or 0,
        hidden_rows=hidden_rows,
        hidden_cols=hidden_cols,
    )

    merged_coords: set[tuple[int, int]] = set()

    for rng in list(ws.merged_cells.ranges):
        raw_span = MergeSpan(
            min_row=rng.min_row,
            min_col=rng.min_col,
            max_row=rng.max_row,
            max_col=rng.max_col,
        )
        origin_val = ws.cell(raw_span.min_row, raw_span.min_col).value

        vis_coords: list[tuple[int, int]] = []
        for r in range(raw_span.min_row, raw_span.max_row + 1):
            for c in range(raw_span.min_col, raw_span.max_col + 1):
                merged_coords.add((r, c))
                if r in hidden_rows or c in hidden_cols:
                    continue
                vis_coords.append((r, c))

        if not vis_coords or (origin_val is None and not propagate_empty):
            continue

        vis_rows = [r for r, _ in vis_coords]
        vis_cols = [c for _, c in vis_coords]
        vis_min_r, vis_max_r = min(vis_rows), max(vis_rows)
        vis_min_c, vis_max_c = min(vis_cols), max(vis_cols)
        vis_span = (
            MergeSpan(
                min_row=vis_min_r,
                min_col=vis_min_c,
                max_row=vis_max_r,
                max_col=vis_max_c,
            )
            if (vis_max_r > vis_min_r or vis_max_c > vis_min_c)
            else None
        )
        if vis_span is not None:
            grid.merges.append(vis_span)

        origin_coord = min(vis_coords)
        for r, c in vis_coords:
            is_origin = (r, c) == origin_coord
            grid.cells[(r, c)] = CellNode(
                row=r,
                col=c,
                raw_value=origin_val if is_origin else None,
                value=origin_val,
                merge=vis_span,
                is_merge_origin=is_origin,
            )

    for row in ws.iter_rows():
        for cell in row:
            r, c = cell.row, cell.column
            if r in hidden_rows or c in hidden_cols:
                continue
            coord = (r, c)
            if coord in merged_coords or cell.value is None:
                continue
            grid.cells[coord] = CellNode(
                row=r,
                col=c,
                raw_value=cell.value,
                value=cell.value,
                merge=None,
                is_merge_origin=True,
            )

    return grid


def unmerge_worksheet_inplace(
    ws: Worksheet,
    *,
    propagate_empty: bool = False,
) -> list[MergeSpan]:
    """Unmerge all ranges directly on `ws` and fill each span with its origin value.

    Returns the list of resolved MergeSpan objects for reference.
    """
    spans: list[MergeSpan] = []

    for rng in list(ws.merged_cells.ranges):
        span = MergeSpan(
            min_row=rng.min_row,
            min_col=rng.min_col,
            max_row=rng.max_row,
            max_col=rng.max_col,
        )
        spans.append(span)
        origin_val = ws.cell(span.min_row, span.min_col).value
        ws.unmerge_cells(str(rng))

        if origin_val is None and not propagate_empty:
            continue

        for r in range(span.min_row, span.max_row + 1):
            for c in range(span.min_col, span.max_col + 1):
                ws.cell(r, c).value = origin_val

    return spans



# Step 3: Formatting feature extraction


# Step 4: Error normalization


