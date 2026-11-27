# Step 2: Connected Component Detection

from collections import deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal
import unicodedata

import openpyxl
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from excel_parsing.preprocess.decompose import CellNode, SheetGrid


@dataclass
class Region:
    region_id: int
    sheet_name: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int
    cells: dict[tuple[int, int], CellNode] = field(default_factory=dict)

    @property
    def num_rows(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def num_cols(self) -> int:
        return self.max_col - self.min_col + 1

    @property
    def area(self) -> int:
        return self.num_rows * self.num_cols

    @property
    def cell_count(self) -> int:
        return len(self.cells)

    @property
    def density(self) -> float:
        return self.cell_count / self.area if self.area > 0 else 0.0

    @property
    def coord_range(self) -> str:
        start = f"{get_column_letter(self.min_col)}{self.min_row}"
        end = f"{get_column_letter(self.max_col)}{self.max_row}"
        return f"{start}:{end}"

    @property
    def kind(self) -> Literal["single_cell", "banner", "vertical_strip", "table"]:
        if self.cell_count == 1:
            return "single_cell"
        if self.num_rows == 1 and self.num_cols > 1:
            return "banner"
        if self.num_cols == 1 and self.num_rows > 1:
            return "vertical_strip"
        return "table"

    def overlaps(self, other: "Region") -> bool:
        return not (
            self.max_row < other.min_row
            or other.max_row < self.min_row
            or self.max_col < other.min_col
            or other.max_col < self.min_col
        )

    def to_matrix(self, fill_value: Any = None) -> list[list[Any]]:
        """Return a dense 2D list of resolved cell values cropped to this region."""
        matrix: list[list[Any]] = []
        for r in range(self.min_row, self.max_row + 1):
            row_vals: list[Any] = []
            for c in range(self.min_col, self.max_col + 1):
                node = self.cells.get((r, c))
                row_vals.append(node.value if node is not None else fill_value)
            matrix.append(row_vals)
        return matrix


def _build_offsets(
    connectivity: Literal[4, 8],
    row_gap: int,
    col_gap: int,
) -> list[tuple[int, int]]:
    if row_gap < 0 or col_gap < 0:
        raise ValueError("row_gap and col_gap must be >= 0")

    max_dr = row_gap + 1
    max_dc = col_gap + 1

    if connectivity == 4:
        offsets: list[tuple[int, int]] = []
        for dr in range(1, max_dr + 1):
            offsets.append((-dr, 0))
            offsets.append((dr, 0))
        for dc in range(1, max_dc + 1):
            offsets.append((0, -dc))
            offsets.append((0, dc))
        return offsets

    if connectivity == 8:
        return [
            (dr, dc)
            for dr in range(-max_dr, max_dr + 1)
            for dc in range(-max_dc, max_dc + 1)
            if not (dr == 0 and dc == 0)
        ]

    raise ValueError(f"connectivity must be 4 or 8, got {connectivity!r}")


def _make_region(
    sheet_name: str,
    cells: dict[tuple[int, int], CellNode],
    region_id: int = 0,
) -> Region:
    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    return Region(
        region_id=region_id,
        sheet_name=sheet_name,
        min_row=min(rows),
        min_col=min(cols),
        max_row=max(rows),
        max_col=max(cols),
        cells=cells,
    )


def _merge_overlapping_regions(regions: list[Region]) -> list[Region]:
    """Iteratively merge regions whose bounding boxes intersect."""
    current = list(regions)
    changed = True

    while changed:
        changed = False
        merged: list[Region] = []
        used = [False] * len(current)

        for i, base in enumerate(current):
            if used[i]:
                continue
            used[i] = True
            combined_cells = dict(base.cells)
            combined = base

            for j in range(i + 1, len(current)):
                if used[j]:
                    continue
                if combined.overlaps(current[j]):
                    used[j] = True
                    combined_cells.update(current[j].cells)
                    combined = _make_region(combined.sheet_name, combined_cells)
                    changed = True

            merged.append(combined)

        current = merged

    return current


_REPEAT_HEADER_ANCHORS: frozenset[str] = frozenset({"no", "no.", "番号", "#"})
_NON_SPLIT_HEADERS: frozenset[str] = frozenset(
    {"単価", "原価", "金額", "数量", "合計", "備考", "容量", "電圧", "種類", "係数", "本数"}
)


def _bfs_components(
    sheet_name: str,
    cells: dict[tuple[int, int], CellNode],
    row_to_vr: dict[int, int],
    col_to_vc: dict[int, int],
    vr_to_row: list[int],
    vc_to_col: list[int],
    offsets: list[tuple[int, int]],
) -> list[Region]:
    v_occupied: dict[tuple[int, int], tuple[int, int]] = {}
    for r, c in cells:
        vr = row_to_vr.get(r)
        vc = col_to_vc.get(c)
        if vr is not None and vc is not None:
            v_occupied[(vr, vc)] = (r, c)

    visited: set[tuple[int, int]] = set()
    out: list[Region] = []

    for start_v in sorted(v_occupied):
        if start_v in visited:
            continue
        queue: deque[tuple[int, int]] = deque([start_v])
        visited.add(start_v)
        comp_cells: dict[tuple[int, int], CellNode] = {}

        while queue:
            vr, vc = queue.popleft()
            orig_coord = v_occupied[(vr, vc)]
            comp_cells[orig_coord] = cells[orig_coord]

            for dvr, dvc in offsets:
                nxt = (vr + dvr, vc + dvc)
                if nxt in v_occupied and nxt not in visited:
                    visited.add(nxt)
                    queue.append(nxt)

        if comp_cells:
            out.append(_make_region(sheet_name, comp_cells))

    return out


def _split_banner_bridged_region(
    region: Region,
    row_to_vr: dict[int, int],
    col_to_vc: dict[int, int],
    vr_to_row: list[int],
    vc_to_col: list[int],
    offsets: list[tuple[int, int]],
) -> list[Region]:
    """Detach wide sheet/section title banners that weld disjoint side-by-side tables."""
    reg_vcs = sorted({col_to_vc[c] for _, c in region.cells if c in col_to_vc})
    if len(reg_vcs) < 6 or region.num_rows < 3:
        return [region]

    min_wide_vcs = max(5, int(len(reg_vcs) * 0.5))
    bridge_rows: set[int] = set()

    for r in sorted({rr for rr, _ in region.cells}):
        row_nodes = [n for (rr, _), n in region.cells.items() if rr == r]
        merges = {n.merge for n in row_nodes if n.merge is not None}
        if len(merges) != 1:
            continue
        m = next(iter(merges))
        if m.row_span > 2:
            continue
        # Every cell on this row inside the region must belong to this single wide merge
        if not all(n.merge == m for n in row_nodes):
            continue
        m_vcs = [vc for c, vc in col_to_vc.items() if m.min_col <= c <= m.max_col]
        if len(m_vcs) >= min_wide_vcs:
            bridge_rows.add(r)

    if not bridge_rows:
        return [region]

    body_cells = {coord: n for coord, n in region.cells.items() if coord[0] not in bridge_rows}
    if not body_cells:
        return [region]

    body_vcs = sorted({col_to_vc[c] for _, c in body_cells if c in col_to_vc})
    if len(body_vcs) < 2:
        return [region]

    # Check whether removing bridge_rows exposes at least one completely empty visible column gap
    has_vc_gap = False
    for vc in range(body_vcs[0] + 1, body_vcs[-1]):
        if vc in body_vcs:
            continue
        col_num = vc_to_col[vc]
        crossed = any(
            n.merge is not None and n.merge.min_col <= col_num <= n.merge.max_col
            for n in body_cells.values()
        )
        if not crossed:
            has_vc_gap = True
            break

    if not has_vc_gap:
        return [region]

    banner_cells = {coord: n for coord, n in region.cells.items() if coord[0] in bridge_rows}
    result: list[Region] = []
    result.extend(
        _bfs_components(region.sheet_name, banner_cells, row_to_vr, col_to_vc, vr_to_row, vc_to_col, offsets)
    )
    result.extend(
        _bfs_components(region.sheet_name, body_cells, row_to_vr, col_to_vc, vr_to_row, vc_to_col, offsets)
    )
    return result


def _split_side_by_side_tables(
    region: Region,
    row_to_vr: dict[int, int],
    col_to_vc: dict[int, int],
    vr_to_row: list[int],
    vc_to_col: list[int],
    offsets: list[tuple[int, int]],
) -> list[Region]:
    """Split adjacent multi-up tables that repeat primary header anchors (e.g. `No | 機器名 | No | 機器名`)."""
    if region.num_cols < 4 or region.num_rows < 3:
        return [region]

    top_r = region.min_row
    top_nodes = sorted(
        (n for (r, _), n in region.cells.items() if r == top_r and n.is_merge_origin),
        key=lambda x: x.col,
    )
    if len(top_nodes) < 2:
        return [region]

    anchor_cols: list[int] = []
    header_scan_rows = [region.min_row, region.min_row + 1]
    for scan_r in header_scan_rows:
        row_origins = sorted(
            (n for (r, _), n in region.cells.items() if r == scan_r and n.is_merge_origin),
            key=lambda x: x.col,
        )
        found = [
            n.col
            for n in row_origins
            if isinstance(n.value, str)
            and unicodedata.normalize("NFKC", n.value).strip().lower() in _REPEAT_HEADER_ANCHORS
        ]
        if len(found) >= 2:
            anchor_cols = found
            top_nodes = row_origins
            break

    candidate_splits: list[int] = []
    if len(anchor_cols) >= 2:
        candidate_splits.extend(anchor_cols[1:])
        last_anchor = anchor_cols[-1]
        for n in top_nodes:
            if n.col < last_anchor + 2 or not isinstance(n.value, str):
                continue
            norm = unicodedata.normalize("NFKC", n.value).strip()
            if norm in _NON_SPLIT_HEADERS or "ベース" in norm:
                continue
            prev_col_count = sum(1 for _, c in region.cells if c == n.col - 1)
            if ("%" in norm or "【" in norm) and prev_col_count <= 2:
                candidate_splits.append(n.col)

    valid_splits = sorted(
        {
            sc
            for sc in candidate_splits
            if region.min_col < sc <= region.max_col
            and not any(
                n.merge is not None and n.merge.min_col < sc <= n.merge.max_col
                for n in region.cells.values()
            )
        }
    )
    if not valid_splits:
        return [region]

    boundaries = [region.min_col] + valid_splits + [region.max_col + 1]
    sub_regions: list[Region] = []
    for idx in range(len(boundaries) - 1):
        c_lo, c_hi = boundaries[idx], boundaries[idx + 1] - 1
        band_cells = {
            coord: node
            for coord, node in region.cells.items()
            if c_lo <= coord[1] <= c_hi
        }
        if band_cells:
            sub_regions.extend(
                _bfs_components(
                    region.sheet_name, band_cells, row_to_vr, col_to_vc, vr_to_row, vc_to_col, offsets
                )
            )
    return sub_regions or [region]


def _is_pure_numeric_region(region: Region) -> bool:
    has_num = False
    for n in region.cells.values():
        if isinstance(n.value, str):
            return False
        if isinstance(n.value, (int, float)) and not isinstance(n.value, bool):
            has_num = True
    return has_num


def _reunite_split_table_columns(
    regions: list[Region],
    col_to_vc: dict[int, int],
) -> list[Region]:
    """Reunite table columns that detached across an empty user-input `数量` column:
    1. Sub-regions whose column range lies within an active header-bearing table above/overlapping them.
    2. Purely numeric price/amount blocks sitting 1-2 visible columns right of an overlapping item-name strip.
    """
    current = sorted(regions, key=lambda r: (r.min_row, r.min_col))
    changed = True

    while changed:
        changed = False
        used = [False] * len(current)
        out: list[Region] = []

        for i, base in enumerate(current):
            if used[i]:
                continue
            used[i] = True
            combined_cells = dict(base.cells)
            combined = base

            for j in range(len(current)):
                if used[j]:
                    continue
                cand = current[j]

                # Rule A: Candidate columns sit inside a single table's header/column span (never re-merge detached sheet-wide banners)
                base_vcs = {col_to_vc[c] for _, c in combined.cells if c in col_to_vc}
                base_rows = {r for r, _ in combined.cells}
                base_merges = {n.merge for n in combined.cells.values() if n.merge is not None}
                is_single_wide_banner = (
                    len(base_rows) <= 2
                    and len(base_merges) == 1
                    and all(n.merge is not None for n in combined.cells.values())
                )
                if (
                    not is_single_wide_banner
                    and 3 <= len(base_vcs) <= 12
                    and cand.min_col >= combined.min_col
                    and cand.max_col <= combined.max_col
                    and cand.min_row >= combined.min_row
                    and cand.min_row <= combined.max_row + 3
                ):
                    used[j] = True
                    combined_cells.update(cand.cells)
                    combined = _make_region(combined.sheet_name, combined_cells)
                    changed = True
                    continue

                # Rule B: Purely numeric orphan price column(s) 1-2 visible cols right of a text region on overlapping rows
                if not _is_pure_numeric_region(combined) and _is_pure_numeric_region(cand):
                    vc_base_max = col_to_vc.get(combined.max_col)
                    vc_cand_min = col_to_vc.get(cand.min_col)
                    if vc_base_max is not None and vc_cand_min is not None:
                        vc_gap = vc_cand_min - vc_base_max
                        ov_start = max(combined.min_row, cand.min_row)
                        ov_end = min(combined.max_row, cand.max_row)
                        overlap_rows = max(0, ov_end - ov_start + 1)
                        min_h = min(combined.num_rows, cand.num_rows)
                        if vc_gap in (1, 2, 3) and min_h > 0 and (overlap_rows / min_h) >= 0.5:
                            used[j] = True
                            combined_cells.update(cand.cells)
                            combined = _make_region(combined.sheet_name, combined_cells)
                            changed = True

            out.append(combined)

        current = sorted(out, key=lambda r: (r.min_row, r.min_col))

    return current


def detect_regions(
    grid: SheetGrid,
    *,
    connectivity: Literal[4, 8] = 4,
    row_gap: int = 0,
    col_gap: int = 0,
    merge_overlapping_bboxes: bool = False,
    split_banner_bridges: bool = True,
    split_side_by_side: bool = True,
    reunite_split_columns: bool = True,
    min_cells: int = 1,
) -> list[Region]:
    """Detect connected spatial regions (tables/blocks) in a resolved SheetGrid.

    Adjacency is evaluated in *visible* row/column coordinates (ignoring hidden rows
    and hidden column ranges) so tables spanning hidden template columns stay intact
    while visible empty separator columns cleanly split neighboring tables.
    """
    if not grid.cells:
        return []

    offsets = _build_offsets(connectivity, row_gap, col_gap)
    max_r = max(grid.max_row, max(r for r, _ in grid.cells))
    max_c = max(grid.max_col, max(c for _, c in grid.cells))

    vr_to_row = [r for r in range(1, max_r + 1) if r not in grid.hidden_rows]
    vc_to_col = [c for c in range(1, max_c + 1) if c not in grid.hidden_cols]
    row_to_vr = {r: idx for idx, r in enumerate(vr_to_row)}
    col_to_vc = {c: idx for idx, c in enumerate(vc_to_col)}

    regions = _bfs_components(
        grid.sheet_name, grid.cells, row_to_vr, col_to_vc, vr_to_row, vc_to_col, offsets
    )

    if split_banner_bridges and regions:
        refined: list[Region] = []
        for reg in regions:
            refined.extend(
                _split_banner_bridged_region(
                    reg, row_to_vr, col_to_vc, vr_to_row, vc_to_col, offsets
                )
            )
        regions = refined

    if split_side_by_side and regions:
        split_out: list[Region] = []
        for reg in regions:
            split_out.extend(
                _split_side_by_side_tables(
                    reg, row_to_vr, col_to_vc, vr_to_row, vc_to_col, offsets
                )
            )
        regions = split_out

    if reunite_split_columns and len(regions) > 1:
        regions = _reunite_split_table_columns(regions, col_to_vc)

    if merge_overlapping_bboxes and len(regions) > 1:
        regions = _merge_overlapping_regions(regions)

    if min_cells > 1:
        regions = [r for r in regions if r.cell_count >= min_cells]

    regions.sort(key=lambda r: (r.min_row, r.min_col))
    for idx, region in enumerate(regions, start=1):
        region.region_id = idx

    return regions


_REGION_PALETTE = [
    "D9E1F2",
    "E2EFDA",
    "FCE4D6",
    "FFF2CC",
    "EDEDED",
    "DDEBF7",
    "F4CCCC",
    "D9D2E9",
    "CFE2F3",
    "D0E0E3",
]


def save_regions_to_excel(
    regions: list[Region],
    output_path: str | Path,
    *,
    include_overlay: bool = True,
    include_region_sheets: bool = True,
) -> Path:
    """Save detected regions into a single Excel workbook for visual inspection.

    Sheets created:
    - `_summary`: index of all detected regions and their stats
    - `_overlay` (if `include_overlay=True`): all regions placed at their original
      `(row, col)` coordinates, color-coded by `region_id` with bounding-box borders
    - `R{id}_{range}` (if `include_region_sheets=True`): each region cropped to A1
    """
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    summary_ws = wb.active
    summary_ws.title = "_summary"

    headers = [
        "region_id",
        "sheet_name",
        "coord_range",
        "kind",
        "rows",
        "cols",
        "cell_count",
        "density",
    ]
    summary_ws.append(headers)
    bold_font = Font(bold=True)
    for col_idx in range(1, len(headers) + 1):
        summary_ws.cell(1, col_idx).font = bold_font

    for r in regions:
        summary_ws.append(
            [
                r.region_id,
                r.sheet_name,
                r.coord_range,
                r.kind,
                r.num_rows,
                r.num_cols,
                r.cell_count,
                round(r.density, 4),
            ]
        )

    if include_overlay:
        overlay_ws = wb.create_sheet("_overlay")
        box_side = Side(style="medium", color="404040")

        for idx, region in enumerate(regions):
            color = _REGION_PALETTE[idx % len(_REGION_PALETTE)]
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

            for r_idx in range(region.min_row, region.max_row + 1):
                for c_idx in range(region.min_col, region.max_col + 1):
                    cell = overlay_ws.cell(r_idx, c_idx)
                    node = region.cells.get((r_idx, c_idx))
                    if node is not None:
                        cell.value = node.value
                    cell.fill = fill
                    cell.border = Border(
                        top=box_side if r_idx == region.min_row else None,
                        bottom=box_side if r_idx == region.max_row else None,
                        left=box_side if c_idx == region.min_col else None,
                        right=box_side if c_idx == region.max_col else None,
                    )

    if include_region_sheets:
        for region in regions:
            safe_range = region.coord_range.replace(":", "_")
            title = f"R{region.region_id:02d}_{safe_range}"[:31]
            ws = wb.create_sheet(title)

            for r_offset, row_vals in enumerate(region.to_matrix(), start=1):
                for c_offset, val in enumerate(row_vals, start=1):
                    if val is not None:
                        ws.cell(r_offset, c_offset).value = val

    wb.save(out_path)
    return out_path

