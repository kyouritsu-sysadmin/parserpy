"""PostgreSQL `json_to_recordset` / `jsonb_to_recordset` query & payload builder."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
import json
from pathlib import Path
import re
from typing import Any, Literal

from excel_parsing.reconcile.reconcile import ImportExtraction, SheetReconciliation

ParamStyle = Literal["dollar", "pyformat", "named"]
JsonFuncType = Literal["json", "jsonb"]
QueryMode = Literal["insert", "select"]

_SIMPLE_IDENT_RE = re.compile(r"^[a-z_][a-z0-9_]*$")
_PG_RESERVED_WORDS: frozenset[str] = frozenset(
    {
        "all",
        "analyse",
        "analyze",
        "and",
        "any",
        "array",
        "as",
        "asc",
        "asymmetric",
        "both",
        "case",
        "cast",
        "check",
        "collate",
        "column",
        "constraint",
        "create",
        "current_catalog",
        "current_date",
        "current_role",
        "current_time",
        "current_timestamp",
        "current_user",
        "default",
        "deferrable",
        "desc",
        "distinct",
        "do",
        "else",
        "end",
        "except",
        "false",
        "fetch",
        "for",
        "foreign",
        "from",
        "grant",
        "group",
        "having",
        "in",
        "initially",
        "intersect",
        "into",
        "lateral",
        "leading",
        "limit",
        "localtime",
        "localtimestamp",
        "not",
        "null",
        "offset",
        "on",
        "only",
        "or",
        "order",
        "placing",
        "primary",
        "references",
        "returning",
        "select",
        "session_user",
        "some",
        "symmetric",
        "table",
        "then",
        "to",
        "trailing",
        "true",
        "union",
        "unique",
        "user",
        "using",
        "variadic",
        "when",
        "where",
        "window",
        "with",
    }
)

_RECORD_FIELD_PG_TYPES: tuple[tuple[str, str], ...] = (
    ("sheet_name", "text"),
    ("source_row", "integer"),
    ("id", "text"),
    ("name", "text"),
    ("category", "text"),
    ("category_rate", "numeric(5,2)"),
    ("group", "text"),
    ("unit_price", "integer"),
    ("note", "text"),
)


def _quote_ident(ident: str) -> str:
    if not ident or "\x00" in ident:
        raise ValueError(f"Invalid SQL identifier: {ident!r}")
    if _SIMPLE_IDENT_RE.match(ident) and ident not in _PG_RESERVED_WORDS:
        return ident
    escaped = ident.replace('"', '""')
    return f'"{escaped}"'


def _quote_table_name(table: str) -> str:
    parts = [p.strip() for p in table.split(".") if p.strip()]
    if not parts or len(parts) > 2:
        raise ValueError(f"Invalid SQL table name: {table!r}")
    return ".".join(_quote_ident(p) for p in parts)


@dataclass(frozen=True)
class PostgresRecordsetPayload:
    """Parameterized PostgreSQL `json_to_recordset` query bundle."""

    query: str
    params: list[str]
    records_json: str
    records: list[dict[str, Any]] = field(default_factory=list)
    row_count: int = 0
    table: str | None = None
    all_accounting_passed: bool = True

    @property
    def named_params(self) -> dict[str, str]:
        return {"records": self.records_json}

    def to_dict(self, *, include_records: bool = False) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "query": self.query,
            "params": list(self.params),
            "row_count": self.row_count,
            "table": self.table,
            "all_accounting_passed": self.all_accounting_passed,
        }
        if include_records:
            payload["records"] = list(self.records)
        return payload

    def to_json(self, *, include_records: bool = False, indent: int = 2) -> str:
        return json.dumps(
            self.to_dict(include_records=include_records),
            ensure_ascii=False,
            indent=indent,
        )

    def save(self, path: str | Path, *, include_records: bool = False) -> Path:
        out = Path(path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(self.to_json(include_records=include_records), encoding="utf-8")
        return out


def denormalize_extraction(
    extraction: ImportExtraction,
    *,
    sheet_name: str | None = None,
    include_sheet_name: bool = False,
) -> list[dict[str, Any]]:
    """Flatten `ImportExtraction` into a single denormalized items list with `category_rate`."""
    rate_by_category = {cat.name: cat.rate for cat in extraction.categories}
    rows: list[dict[str, Any]] = []

    for item in extraction.items:
        row: dict[str, Any] = {}
        if include_sheet_name:
            row["sheet_name"] = sheet_name
        row.update(
            {
                "source_row": item.source_row,
                "id": item.id,
                "name": item.name,
                "category": item.category,
                "category_rate": rate_by_category.get(item.category),
                "group": item.group,
                "unit_price": item.unit_price,
                "note": item.note,
            }
        )
        rows.append(row)

    return rows


def build_pg_recordset_payload(
    extraction_or_sheets: ImportExtraction | SheetReconciliation | Sequence[SheetReconciliation],
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
    """Build a parameterized PostgreSQL `json_to_recordset` query + JSON array payload.

    Args:
        extraction_or_sheets: A validated `ImportExtraction`, a single `SheetReconciliation`,
            or a sequence of `SheetReconciliation`s (multi-sheet workbook).
        table: Target PostgreSQL table name (used when `mode="insert"`).
        mode: `"insert"` (`INSERT INTO <table> ... SELECT ... FROM json_to_recordset(...)`)
            or `"select"` (`SELECT ... FROM json_to_recordset(...)`).
        include_sheet_name: Include `sheet_name` column in the recordset. Defaults to `True`
            when multiple sheets are passed, else `False`.
        include_fields: Optional subset of record fields to project into the SQL statement.
        column_map: Optional mapping from record field name to target DB table column name
            (e.g. `{"id": "catalog_id", "group": "item_group"}`).
        extra_columns: Optional mapping of extra target column names to SQL bind expressions
            prepended before recordset columns (e.g. `{"import_batch_id": "$2::uuid"}`).
        param_style: `"dollar"` (`$1::json` for `asyncpg`), `"pyformat"` (`%s::json` for
            `psycopg`), or `"named"` (`CAST(:records AS json)` for SQLAlchemy `text()`).
        json_type: `"json"` (`json_to_recordset`) or `"jsonb"` (`jsonb_to_recordset`).
        on_conflict: Optional `ON CONFLICT` clause body (without leading `"ON CONFLICT"`),
            e.g. `"(id) DO UPDATE SET unit_price = EXCLUDED.unit_price"`.
        returning: Optional `RETURNING` clause body (without leading `"RETURNING"`).
        alias: Table alias for the `json_to_recordset` lateral source (default `"r"`).
    """
    all_passed = True
    records: list[dict[str, Any]] = []

    if isinstance(extraction_or_sheets, ImportExtraction):
        use_sheet_col = bool(include_sheet_name)
        records = denormalize_extraction(
            extraction_or_sheets,
            sheet_name=None,
            include_sheet_name=use_sheet_col,
        )
    elif isinstance(extraction_or_sheets, SheetReconciliation):
        use_sheet_col = bool(include_sheet_name)
        all_passed = extraction_or_sheets.all_accounting_passed and (
            extraction_or_sheets.extraction is not None
        )
        if extraction_or_sheets.extraction is not None:
            records = denormalize_extraction(
                extraction_or_sheets.extraction,
                sheet_name=extraction_or_sheets.sheet_name,
                include_sheet_name=use_sheet_col,
            )
    else:
        sheets = list(extraction_or_sheets)
        use_sheet_col = (len(sheets) > 1) if include_sheet_name is None else include_sheet_name
        all_passed = all(
            s.all_accounting_passed and s.extraction is not None for s in sheets
        )
        for s in sheets:
            if s.extraction is not None:
                records.extend(
                    denormalize_extraction(
                        s.extraction,
                        sheet_name=s.sheet_name,
                        include_sheet_name=use_sheet_col,
                    )
                )

    available_defs = [
        (fname, pg_type)
        for fname, pg_type in _RECORD_FIELD_PG_TYPES
        if fname != "sheet_name" or use_sheet_col
    ]

    if include_fields is not None:
        allowed = set(include_fields)
        unknown = allowed - {f for f, _ in available_defs}
        if unknown:
            raise ValueError(f"Unknown recordset field(s): {sorted(unknown)}")
        active_defs = [(f, t) for f, t in available_defs if f in allowed]
        if not active_defs:
            raise ValueError("include_fields cannot be empty")
        records = [{k: row[k] for k, _ in active_defs} for row in records]
    else:
        active_defs = available_defs

    col_aliases = dict(column_map or {})
    extra_cols = dict(extra_columns or {})
    alias_ident = _quote_ident(alias)

    fn_name = "jsonb_to_recordset" if json_type == "jsonb" else "json_to_recordset"
    if param_style == "dollar":
        bind_expr = f"$1::{json_type}"
    elif param_style == "pyformat":
        bind_expr = f"%s::{json_type}"
    elif param_style == "named":
        bind_expr = f"CAST(:records AS {json_type})"
    else:
        raise ValueError(f"Unsupported param_style: {param_style!r}")

    recordset_sig = ", ".join(f"{_quote_ident(fname)} {pg_type}" for fname, pg_type in active_defs)

    target_cols: list[str] = []
    select_exprs: list[str] = []

    for extra_col, sql_expr in extra_cols.items():
        target_cols.append(_quote_ident(extra_col))
        if mode == "select":
            select_exprs.append(f"{sql_expr} AS {_quote_ident(extra_col)}")
        else:
            select_exprs.append(sql_expr)

    for fname, _ in active_defs:
        src_ref = f"{alias_ident}.{_quote_ident(fname)}"
        dest_col = col_aliases.get(fname, fname)
        target_cols.append(_quote_ident(dest_col))
        if mode == "select" and dest_col != fname:
            select_exprs.append(f"{src_ref} AS {_quote_ident(dest_col)}")
        else:
            select_exprs.append(src_ref)

    from_clause = f"FROM {fn_name}({bind_expr}) AS {alias_ident}({recordset_sig})"

    if mode == "select":
        query = f"SELECT {', '.join(select_exprs)} {from_clause}"
        effective_table: str | None = None
    elif mode == "insert":
        table_ident = _quote_table_name(table)
        parts = [
            f"INSERT INTO {table_ident} ({', '.join(target_cols)})",
            f"SELECT {', '.join(select_exprs)}",
            from_clause,
        ]
        if on_conflict:
            cleaned_conflict = on_conflict.strip()
            if cleaned_conflict.upper().startswith("ON CONFLICT"):
                parts.append(cleaned_conflict)
            else:
                parts.append(f"ON CONFLICT {cleaned_conflict}")
        if returning:
            cleaned_ret = returning.strip()
            if cleaned_ret.upper().startswith("RETURNING"):
                parts.append(cleaned_ret)
            else:
                parts.append(f"RETURNING {cleaned_ret}")
        query = " ".join(parts)
        effective_table = table
    else:
        raise ValueError(f"Unsupported mode: {mode!r}")

    records_json = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    return PostgresRecordsetPayload(
        query=query,
        params=[records_json],
        records_json=records_json,
        records=records,
        row_count=len(records),
        table=effective_table,
        all_accounting_passed=all_passed,
    )
