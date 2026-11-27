# Compact semantic HTML serialization for LLM parsing

import html
from pathlib import Path
from typing import Any

from openpyxl.utils import get_column_letter

from excel_parsing.preprocess.decompose import CellNode
from excel_parsing.preprocess.detection import Region
from excel_parsing.preprocess.formatting import CellFormat, RowRole, classify_region_rows


def _format_cell_value(
    value: Any,
    fmt: CellFormat | None,
    *,
    format_percents: bool = True,
) -> str:
    if value is None:
        return ""

    if (
        format_percents
        and fmt is not None
        and fmt.format_type == "percent"
        and isinstance(value, (int, float))
        and not isinstance(value, bool)
    ):
        pct = value * 100
        text = f"{pct:.4f}".rstrip("0").rstrip(".") + "%"
    else:
        text = str(value)

    escaped = html.escape(text, quote=False)
    return escaped.replace("\n", "<br>")


def _build_cell_attrs(
    r: int,
    c: int,
    node: CellNode | None,
    fmt: CellFormat | None,
    row_span: int,
    col_span: int,
    *,
    include_coords: bool,
    include_empty_coords: bool,
) -> str:
    attrs: list[str] = []

    if include_coords and (node is not None or include_empty_coords):
        if node is not None and node.merge is not None and (row_span > 1 or col_span > 1):
            coord = node.merge.coord_range
        else:
            coord = f"{get_column_letter(c)}{r}"
        attrs.append(f'c="{coord}"')

    if row_span > 1:
        attrs.append(f'rowspan="{row_span}"')
    if col_span > 1:
        attrs.append(f'colspan="{col_span}"')

    if fmt is not None and fmt.indent > 0:
        attrs.append(f'indent="{fmt.indent}"')

    return (" " + " ".join(attrs)) if attrs else ""


def region_to_html(
    region: Region,
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
    *,
    rows: list[int] | None = None,
    include_coords: bool = True,
    include_empty_coords: bool = False,
    format_percents: bool = True,
    drop_empty_cols: bool = True,
) -> str:
    """Serialize a single `Region` into compact, LLM-friendly native HTML.

    - Uses native `rowspan` / `colspan` computed over rendered visible rows/columns
    - Drops completely empty/hidden columns (`drop_empty_cols=True`) and empty rows
      so sparse or wide sheets do not emit thousands of empty `<td></td>` tags
    - Supports an explicit `rows` subset (e.g. non-contiguous retry rows) without
      fabricating blank intermediate `<tr>` rows
    - Adds `c="A1"` (or `c="A1:C1"` for merged blocks) cell coordinates when `include_coords=True`
    - Groups leading header/section rows into `<thead>` and body/total rows into `<tbody>`
    """
    fmts = formats or {}
    roles = row_roles if row_roles is not None else classify_region_rows(region, fmts)

    if rows is not None:
        rendered_rows = sorted({r for r in rows if region.min_row <= r <= region.max_row})
    elif drop_empty_cols:
        rendered_rows = sorted({r for r, _ in region.cells})
    else:
        rendered_rows = list(range(region.min_row, region.max_row + 1))

    if not rendered_rows:
        return (
            f'<table data-region="{region.region_id}" '
            f'data-range="{region.coord_range}" data-kind="{region.kind}"></table>'
        )

    rendered_row_set = set(rendered_rows)
    if drop_empty_cols:
        rendered_cols = sorted({c for r, c in region.cells if r in rendered_row_set})
        if not rendered_cols:
            rendered_cols = sorted({c for _, c in region.cells})
    else:
        rendered_cols = list(range(region.min_col, region.max_col + 1))

    # Identify leading contiguous header/section band for <thead>
    thead_rows: set[int] = set()
    for r in rendered_rows:
        if roles.get(r, "data") in ("header", "section"):
            thead_rows.add(r)
        else:
            break

    # If the entire rendered slice is headers/sections, keep it in <tbody>
    if len(thead_rows) == len(rendered_rows):
        thead_rows.clear()

    covered: set[tuple[int, int]] = set()
    head_lines: list[str] = []
    body_lines: list[str] = []

    for r in rendered_rows:
        role = roles.get(r, "data")
        in_thead = r in thead_rows
        tag = "th" if role in ("header", "section") else "td"

        tr_attr = ""
        if role == "section":
            tr_attr = ' role="section"'
        elif role == "total":
            tr_attr = ' role="total"'
        elif role == "header" and not in_thead:
            tr_attr = ' role="header"'

        cells_html: list[str] = []
        for c in rendered_cols:
            if (r, c) in covered:
                continue

            node = region.cells.get((r, c))
            fmt = fmts.get((r, c)) if node is not None else None
            row_span = 1
            col_span = 1
            coord_override: str | None = None

            if node is not None and node.merge is not None:
                span = node.merge
                spanned_rows = [
                    rr
                    for rr in rendered_rows
                    if max(r, span.min_row) <= rr <= min(region.max_row, span.max_row)
                    and ((rr in thead_rows) == in_thead)
                ]
                spanned_cols = [
                    cc
                    for cc in rendered_cols
                    if max(c, span.min_col) <= cc <= min(region.max_col, span.max_col)
                ]
                row_span = max(1, len(spanned_rows))
                col_span = max(1, len(spanned_cols))
                for rr in spanned_rows:
                    for cc in spanned_cols:
                        if (rr, cc) != (r, c):
                            covered.add((rr, cc))
                if row_span > 1 or col_span > 1:
                    c_start = f"{get_column_letter(spanned_cols[0])}{spanned_rows[0]}"
                    c_end = f"{get_column_letter(spanned_cols[-1])}{spanned_rows[-1]}"
                    coord_override = f"{c_start}:{c_end}"
                else:
                    coord_override = f"{get_column_letter(c)}{r}"

            if coord_override is not None and include_coords:
                attrs: list[str] = [f'c="{coord_override}"']
                if row_span > 1:
                    attrs.append(f'rowspan="{row_span}"')
                if col_span > 1:
                    attrs.append(f'colspan="{col_span}"')
                if fmt is not None and fmt.indent > 0:
                    attrs.append(f'indent="{fmt.indent}"')
                attr_str = " " + " ".join(attrs)
            else:
                attr_str = _build_cell_attrs(
                    r,
                    c,
                    node,
                    fmt,
                    row_span,
                    col_span,
                    include_coords=include_coords,
                    include_empty_coords=include_empty_coords,
                )

            content = (
                _format_cell_value(node.value, fmt, format_percents=format_percents)
                if node is not None
                else ""
            )
            cells_html.append(f"<{tag}{attr_str}>{content}</{tag}>")

        row_line = f"<tr{tr_attr}>{''.join(cells_html)}</tr>"
        if in_thead:
            head_lines.append(row_line)
        else:
            body_lines.append(row_line)

    parts = [
        f'<table data-region="{region.region_id}" data-range="{region.coord_range}" data-kind="{region.kind}">'
    ]
    if head_lines:
        parts.append("<thead>")
        parts.extend(head_lines)
        parts.append("</thead>")
    if body_lines:
        parts.append("<tbody>")
        parts.extend(body_lines)
        parts.append("</tbody>")
    parts.append("</table>")

    return "\n".join(parts)


def sheet_to_html(
    sheet_name: str,
    regions: list[Region],
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles_by_region: dict[int, dict[int, RowRole]] | None = None,
    *,
    include_coords: bool = True,
    include_empty_coords: bool = False,
    format_percents: bool = True,
) -> str:
    """Serialize all regions of a sheet into one compact HTML `<section>`."""
    escaped_name = html.escape(sheet_name, quote=True)
    tables = [
        region_to_html(
            region,
            formats=formats,
            row_roles=(
                row_roles_by_region.get(region.region_id)
                if row_roles_by_region is not None
                else None
            ),
            include_coords=include_coords,
            include_empty_coords=include_empty_coords,
            format_percents=format_percents,
        )
        for region in regions
    ]
    body = "\n".join(tables)
    return f'<section data-sheet="{escaped_name}">\n{body}\n</section>'


def save_html(
    html_content: str,
    output_path: str | Path,
    *,
    standalone: bool = False,
    title: str = "Sheet Preview",
) -> Path:
    """Write HTML to disk.

    - `standalone=False` (default): writes the exact compact LLM-ready HTML fragment.
    - `standalone=True`: wraps with minimal `<!DOCTYPE html>` + table borders for browser viewing.
    """
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if standalone:
        doc = (
            "<!DOCTYPE html>\n"
            '<html lang="en">\n<head>\n<meta charset="utf-8">\n'
            f"<title>{html.escape(title)}</title>\n"
            "<style>\n"
            "body{font-family:sans-serif;font-size:13px;margin:16px}\n"
            "table{border-collapse:collapse;margin:16px 0}\n"
            "th,td{border:1px solid #ccc;padding:4px 8px;vertical-align:top}\n"
            "th{background:#eef4fb;font-weight:600}\n"
            'tr[role="section"] th{background:#e8e0f5;text-align:left}\n'
            'tr[role="total"] td{background:#fff6d6;font-weight:600}\n'
            "</style>\n</head>\n<body>\n"
            f"{html_content}\n"
            "</body>\n</html>\n"
        )
        out_path.write_text(doc, encoding="utf-8")
    else:
        out_path.write_text(html_content, encoding="utf-8")

    return out_path
