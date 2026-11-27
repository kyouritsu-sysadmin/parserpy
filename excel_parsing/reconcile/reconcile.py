# Stage 4: Reconciliation (Row-Accounting Invariant + Two-Tier Hard Schema Validation)

from dataclasses import asdict, dataclass, field
import json
from pathlib import Path
import re
from typing import Annotated, Any, Literal
import unicodedata

from pydantic import BaseModel, BeforeValidator, ConfigDict, Field, ValidationError, field_validator

try:
    from database.schemas.equipment_catalog_schema import CATALOG_MAX_UNIT_PRICE, CatalogId
except ImportError:
    CATALOG_MAX_UNIT_PRICE = 100_000_000
    CatalogId = Annotated[str, Field(min_length=1, max_length=64)]

from excel_parsing.extraction.extract import (
    ExtractedRecord,
    ExtractionResult,
    ParentState,
    RowClassification,
    StatefulChunk,
    VLLMClient,
    advance_parent_state,
    extract_chunk_records,
)
from excel_parsing.preprocess.detection import Region
from excel_parsing.preprocess.formatting import CellFormat, RowRole
from excel_parsing.serialize.serialize import region_to_html


# ---------------------------------------------------------------------------
# Tier-2 Strict Domain Schema (Pydantic)
# ---------------------------------------------------------------------------

MAX_IMPORT_ITEMS = 2000


def clean_text(value: Any) -> str:
    if not isinstance(value, str):
        raise ValueError("Expected text")
    if any(unicodedata.category(c) == "Cc" and c not in "\r\n\t" for c in value):
        raise ValueError("Control characters are not allowed")
    return " ".join(value.split())


Text = Annotated[str, BeforeValidator(clean_text), Field(min_length=1, max_length=255)]
Name = Annotated[str, BeforeValidator(clean_text), Field(min_length=1, max_length=2000)]
Row = Annotated[int, Field(strict=True, ge=1, le=20_000)]


class ImportModel(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True, frozen=True)


class ImportCategory(ImportModel):
    name: Text
    rate: Annotated[float, Field(ge=0, le=100, allow_inf_nan=False)] | None

    @field_validator("rate", mode="before")
    @classmethod
    def strict_rate(cls, value: Any) -> float | None:
        if isinstance(value, bool) or (value is not None and not isinstance(value, (int, float))):
            raise ValueError("Rate must be a number or null")
        if value is not None:
            fval = float(value)
            if abs(fval - round(fval, 2)) > 1e-9:
                raise ValueError("Rate supports at most two decimal places")
            return round(fval, 2)
        return None


class ImportItem(ImportModel):
    source_row: Row
    id: CatalogId | None
    name: Name
    category: Text
    group: Text | None
    unit_price: Annotated[int, Field(strict=True, ge=0, le=CATALOG_MAX_UNIT_PRICE)] | None
    note: Text | None

    @field_validator("id", mode="before")
    @classmethod
    def normalize_id(cls, value: Any) -> Any:
        return unicodedata.normalize("NFKC", value).strip().upper() if isinstance(value, str) else value


class IgnoredRow(ImportModel):
    source_row: Row
    reason: Literal["heading", "separator", "description", "total"]


class ImportExtraction(ImportModel):
    categories: Annotated[list[ImportCategory], Field(max_length=100)]
    items: Annotated[list[ImportItem], Field(max_length=MAX_IMPORT_ITEMS)]
    ignored_rows: Annotated[list[IgnoredRow], Field(max_length=MAX_IMPORT_ITEMS)]


# ---------------------------------------------------------------------------
# Reconciliation Diagnostics
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RowAccountingResult:
    region_id: int
    input_row_count: int
    accounted_row_count: int
    missing_rows: list[int] = field(default_factory=list)
    hallucinated_rows: list[int] = field(default_factory=list)
    duplicate_item_rows: list[int] = field(default_factory=list)
    overlap_rows: list[int] = field(default_factory=list)
    invalid_item_rows: list[int] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return not (
            self.missing_rows
            or self.hallucinated_rows
            or self.overlap_rows
            or self.duplicate_item_rows
            or self.invalid_item_rows
        )


@dataclass(frozen=True)
class RowValidationIssue:
    region_id: int
    source_row: int | None
    model: str
    field: str
    error: str
    raw_value: Any = None


@dataclass
class RegionReconciliation:
    region_id: int
    accounting: RowAccountingResult
    categories: list[ImportCategory] = field(default_factory=list)
    items: list[ImportItem] = field(default_factory=list)
    ignored_rows: list[IgnoredRow] = field(default_factory=list)
    validation_issues: list[RowValidationIssue] = field(default_factory=list)
    retried_rows: list[int] = field(default_factory=list)


@dataclass
class SheetReconciliation:
    sheet_name: str
    regions: list[RegionReconciliation] = field(default_factory=list)
    extraction: ImportExtraction | None = None
    sheet_errors: list[str] = field(default_factory=list)

    @property
    def all_accounting_passed(self) -> bool:
        return all(r.accounting.passed for r in self.regions)

    def to_report_dict(self) -> dict[str, Any]:
        return {
            "sheet_name": self.sheet_name,
            "all_accounting_passed": self.all_accounting_passed,
            "total_categories": len(self.extraction.categories) if self.extraction else 0,
            "total_valid_items": sum(len(r.items) for r in self.regions),
            "total_ignored_rows": sum(len(r.ignored_rows) for r in self.regions),
            "total_validation_issues": sum(len(r.validation_issues) for r in self.regions),
            "sheet_errors": self.sheet_errors,
            "regions": [
                {
                    "region_id": r.region_id,
                    "accounting": {
                        **asdict(r.accounting),
                        "passed": r.accounting.passed,
                    },
                    "valid_items": len(r.items),
                    "ignored_rows": len(r.ignored_rows),
                    "retried_rows": r.retried_rows,
                    "validation_issues": [asdict(i) for i in r.validation_issues],
                }
                for r in self.regions
            ],
        }


# ---------------------------------------------------------------------------
# Soft -> Hard Field Adapters
# ---------------------------------------------------------------------------

_ROLE_TO_IGNORED_REASON: dict[str, Literal["heading", "separator", "description", "total"]] = {
    "category_header": "heading",
    "group_header": "heading",
    "column_header": "heading",
    "subtotal": "total",
    "footnote": "description",
    "other": "description",
}

_ID_EXACT_KEYS = ("id", "catalog_id", "no", "no.", "番号", "品番", "型式", "コード", "管理番号")
_PRICE_EXACT_KEYS = (
    "unit_price",
    "単価",
    "原価",
    "銅バー単価",
    "価格",
    "見積価格",
    "標準単価",
    "定価",
    "price",
)
_PRICE_SUBSTR_KEYS = ("単価", "原価", "ベース", "unit_price")
_NOTE_EXACT_KEYS = ("note", "notes", "備考", "注記", "仕様", "コメント", "摘要")
_DASH_NULL_TOKENS: frozenset[str] = frozenset({"-", "ー", "―", "‐", "–", "—", "なし", "無し", "n/a", "na"})

_PCT_RATE_RE = re.compile(r"(?<![A-Za-z0-9.])(\d{1,3}(?:\.\d{1,2})?)\s*%")
_TRAILING_RATE_RE = re.compile(r"\s+(\d{1,3}(?:\.\d{1,2})?)\s*$")


def _norm_key(key: str) -> str:
    collapsed = " ".join(unicodedata.normalize("NFKC", str(key)).split())
    return collapsed.lower()


def _normalize_fields(fields: dict[str, Any]) -> list[tuple[str, str, Any]]:
    """Return `(orig_key, normalized_lower_key, value)` for non-empty field entries."""
    out: list[tuple[str, str, Any]] = []
    for k, v in fields.items():
        if v is None or v == "":
            continue
        out.append((str(k), _norm_key(str(k)), v))
    return out


def _extract_id_from_fields(fields: dict[str, Any]) -> Any:
    entries = _normalize_fields(fields)
    for _, nk, v in entries:
        if nk in _ID_EXACT_KEYS:
            return v
    return None


def _extract_price_from_fields(fields: dict[str, Any]) -> Any:
    entries = _normalize_fields(fields)
    # 1. Exact priority match
    for target in _PRICE_EXACT_KEYS:
        for _, nk, v in entries:
            if nk == target:
                return v
    # 2. Substring match on normalized header (e.g. "2026.1ベース", "銅バー単価")
    for _, nk, v in entries:
        if any(sub in nk for sub in _PRICE_SUBSTR_KEYS):
            return v
    # 3. Fallback to "金額" only if positive numeric (never copy `[ERROR:...]` from computed amount)
    for _, nk, v in entries:
        if (
            nk == "金額"
            and isinstance(v, (int, float))
            and not isinstance(v, bool)
            and v > 0
        ):
            return v
    # 4. Headerless 2-column table fallback (single generic integer-valued numeric field)
    generic_nums = [
        v
        for _, nk, v in entries
        if nk not in _ID_EXACT_KEYS
        and nk not in _NOTE_EXACT_KEYS
        and nk not in ("数量", "数量 合計", "合計", "金額", "容量", "電圧", "本数", "係数", "薄型係数")
        and not isinstance(v, bool)
        and (
            isinstance(v, int)
            or (isinstance(v, float) and abs(v - round(v)) < 1e-9)
        )
    ]
    if len(generic_nums) == 1:
        return generic_nums[0]
    return None


def _extract_note_from_fields(fields: dict[str, Any]) -> Any:
    entries = _normalize_fields(fields)
    for _, nk, v in entries:
        if nk in _NOTE_EXACT_KEYS:
            return v
    return None


def _parse_category_rate(label: str) -> float | None:
    """Extract an optional category rate (0..100) from section banners like
    `'特殊機器 40% (ACB/IGR/MDU)'` or `'ACB(ZCT付) 40'`.
    """
    norm = " ".join(unicodedata.normalize("NFKC", label).split())
    if not norm or norm.lower().startswith("row "):
        return None

    m_pct = _PCT_RATE_RE.search(norm)
    if m_pct:
        val = float(m_pct.group(1))
        if 0.0 <= val <= 100.0:
            return round(val, 2)

    m_tail = _TRAILING_RATE_RE.search(norm)
    if m_tail:
        val = float(m_tail.group(1))
        if 0.0 < val <= 100.0:
            return round(val, 2)

    return None


def _coerce_soft_int(value: Any) -> Any:
    """Safely coerce exact integer floats, currency strings, or placeholder dashes for Tier-2."""
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(round(value)) if abs(value - round(value)) < 1e-9 else value
    if isinstance(value, str):
        norm = unicodedata.normalize("NFKC", value).strip()
        if not norm or norm.lower() in _DASH_NULL_TOKENS:
            return None
        if norm.startswith("[ERROR:"):
            return value
        cleaned = norm.replace(",", "").replace("¥", "").replace("￥", "").replace("円", "").strip()
        if cleaned.lower() in _DASH_NULL_TOKENS:
            return None
        try:
            fval = float(cleaned)
            return int(round(fval)) if abs(fval - round(fval)) < 1e-9 else value
        except ValueError:
            return value
    return value


def _coerce_soft_str(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return str(value)


def check_row_accounting(
    region_id: int,
    input_rows: set[int],
    item_rows: list[int],
    ignored_rows: list[int],
    *,
    invalid_item_rows: list[int] | None = None,
) -> RowAccountingResult:
    """Verify R_input == R_items ∪ R_ignored with zero hallucinations, duplicates, overlaps, or drops."""
    item_set = set(item_rows)
    ignored_set = set(ignored_rows)
    accounted = item_set | ignored_set

    seen: set[int] = set()
    dups: list[int] = []
    for r in item_rows:
        if r in seen and r not in dups:
            dups.append(r)
        seen.add(r)

    return RowAccountingResult(
        region_id=region_id,
        input_row_count=len(input_rows),
        accounted_row_count=len(accounted & input_rows),
        missing_rows=sorted(input_rows - accounted),
        hallucinated_rows=sorted(accounted - input_rows),
        duplicate_item_rows=sorted(dups),
        overlap_rows=sorted(item_set & ignored_set),
        invalid_item_rows=sorted(set(invalid_item_rows or [])),
    )


def _retry_missing_rows(
    client: VLLMClient,
    region: Region,
    skeleton: list[RowClassification],
    missing_rows: list[int],
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
) -> list[ExtractedRecord]:
    """Re-extract only the missing rows as a targeted sub-chunk without blank gap rows."""
    skeleton_by_row = {rc.row: rc for rc in skeleton}
    ordered_missing = sorted(set(missing_rows))
    min_miss = ordered_missing[0]

    prior_rows = [rc.row for rc in skeleton if rc.row < min_miss]
    parent_state = advance_parent_state(
        ParentState(),
        prior_rows,
        skeleton_by_row,
        region,
        formats=formats,
        include_coords=True,
    )

    miss_set = set(ordered_missing)
    sub_cells = {
        coord: node for coord, node in region.cells.items() if coord[0] in miss_set
    }
    cols = [c for _, c in sub_cells] if sub_cells else [region.min_col, region.max_col]
    sub_region = Region(
        region_id=region.region_id,
        sheet_name=region.sheet_name,
        min_row=ordered_missing[0],
        min_col=min(cols),
        max_row=ordered_missing[-1],
        max_col=max(cols),
        cells=sub_cells,
    )

    chunk = StatefulChunk(
        chunk_id=999,
        region_id=region.region_id,
        start_row=ordered_missing[0],
        end_row=ordered_missing[-1],
        rows=ordered_missing,
        parent_state=parent_state,
        html=region_to_html(
            sub_region,
            formats=formats,
            row_roles=row_roles,
            rows=ordered_missing,
            include_coords=True,
        ),
    )
    return extract_chunk_records(client, chunk, skeleton_by_row, region=region)


def _deduplicate_records_by_row(records: list[ExtractedRecord]) -> list[ExtractedRecord]:
    """Keep one record per `source_row`, preferring the one with non-empty name and richer fields."""
    by_row: dict[int, ExtractedRecord] = {}
    order: list[int] = []
    for rec in records:
        r = rec.source_row
        if r not in by_row:
            by_row[r] = rec
            order.append(r)
        else:
            prev = by_row[r]
            if (not prev.name and rec.name) or (len(rec.fields) > len(prev.fields)):
                by_row[r] = rec
    return [by_row[r] for r in order]


def reconcile_region_data(
    region_id: int,
    skeleton: list[RowClassification],
    records: list[ExtractedRecord],
    *,
    region: Region | None = None,
    client: VLLMClient | None = None,
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
    retry_missing: bool = False,
) -> RegionReconciliation:
    """Run Row-Accounting Invariant + per-record Tier-2 Pydantic validation on one region."""
    input_rows = {rc.row for rc in skeleton}
    skeleton_by_row = {rc.row: rc for rc in skeleton}

    working_records = list(records)
    retried_rows: list[int] = []

    def _partition(recs: list[ExtractedRecord]) -> tuple[list[ExtractedRecord], list[IgnoredRow]]:
        # Filter out non-line_item records or hallucinated header records when skeleton says non-line_item
        raw_line_recs = [
            r
            for r in recs
            if r.record_type == "line_item"
            and (r.source_row not in skeleton_by_row or skeleton_by_row[r.source_row].role == "line_item")
        ]
        line_recs = _deduplicate_records_by_row(raw_line_recs)
        line_row_set = {r.source_row for r in line_recs}
        ignored: list[IgnoredRow] = []
        for rc in skeleton:
            if rc.role == "line_item" or rc.row in line_row_set:
                continue
            reason = (
                "separator"
                if rc.role == "other" and not rc.label
                else _ROLE_TO_IGNORED_REASON.get(rc.role, "description")
            )
            ignored.append(IgnoredRow(source_row=rc.row, reason=reason))
        return line_recs, ignored

    line_recs, ignored_rows = _partition(working_records)
    accounting = check_row_accounting(
        region_id,
        input_rows,
        [r.source_row for r in line_recs],
        [ig.source_row for ig in ignored_rows],
    )

    # Optional targeted retry for missing line_item rows only
    if retry_missing and accounting.missing_rows and client is not None and region is not None:
        retry_targets = [
            r for r in accounting.missing_rows if skeleton_by_row.get(r) and skeleton_by_row[r].role == "line_item"
        ]
        if retry_targets:
            retried_rows = retry_targets
            retry_recs = _retry_missing_rows(
                client, region, skeleton, retry_targets, formats=formats, row_roles=row_roles
            )
            working_records.extend(retry_recs)
            records.extend(retry_recs)
            line_recs, ignored_rows = _partition(working_records)

    issues: list[RowValidationIssue] = []
    valid_items: list[ImportItem] = []
    invalid_rows: list[int] = []
    seen_categories: dict[str, ImportCategory] = {}

    # Validate categories discovered in skeleton / records (including parsed rate)
    cat_names: list[str] = [
        rc.label for rc in skeleton if rc.role == "category_header" and rc.label
    ]
    for rec in line_recs:
        if rec.category and rec.category not in cat_names:
            cat_names.append(rec.category)

    for cname in cat_names:
        cleaned_cname = " ".join(cname.split())
        if not cleaned_cname or cleaned_cname in seen_categories:
            continue
        parsed_rate = _parse_category_rate(cleaned_cname)
        try:
            seen_categories[cleaned_cname] = ImportCategory(name=cleaned_cname, rate=parsed_rate)
        except ValidationError as exc:
            for err in exc.errors():
                issues.append(
                    RowValidationIssue(
                        region_id=region_id,
                        source_row=None,
                        model="ImportCategory",
                        field=".".join(str(p) for p in err["loc"]),
                        error=err["msg"],
                        raw_value=cname,
                    )
                )

    # Validate each line_item record individually against ImportItem
    for rec in line_recs:
        raw_id = _coerce_soft_str(_extract_id_from_fields(rec.fields))
        raw_price = _coerce_soft_int(_extract_price_from_fields(rec.fields))
        raw_note = _coerce_soft_str(_extract_note_from_fields(rec.fields))
        raw_cat = rec.category or "Uncategorized"
        raw_group = rec.group if rec.group and rec.group != "Default" else None
        raw_name = rec.name or (skeleton_by_row[rec.source_row].label if rec.source_row in skeleton_by_row else "")

        payload = {
            "source_row": rec.source_row,
            "id": raw_id,
            "name": raw_name,
            "category": raw_cat,
            "group": raw_group,
            "unit_price": raw_price,
            "note": raw_note,
        }
        try:
            item = ImportItem.model_validate(payload)
            valid_items.append(item)
            if item.category not in seen_categories:
                seen_categories[item.category] = ImportCategory(
                    name=item.category,
                    rate=_parse_category_rate(item.category),
                )
        except ValidationError as exc:
            invalid_rows.append(rec.source_row)
            for err in exc.errors():
                field_name = ".".join(str(p) for p in err["loc"])
                issues.append(
                    RowValidationIssue(
                        region_id=region_id,
                        source_row=rec.source_row,
                        model="ImportItem",
                        field=field_name,
                        error=err["msg"],
                        raw_value=payload.get(field_name),
                    )
                )

    # Final accounting reflects validated items + ignored rows (invalid items surface explicitly)
    accounting = check_row_accounting(
        region_id,
        input_rows,
        [it.source_row for it in valid_items],
        [ig.source_row for ig in ignored_rows],
        invalid_item_rows=invalid_rows,
    )

    return RegionReconciliation(
        region_id=region_id,
        accounting=accounting,
        categories=list(seen_categories.values()),
        items=valid_items,
        ignored_rows=ignored_rows,
        validation_issues=issues,
        retried_rows=retried_rows,
    )


def reconcile_region(
    result: ExtractionResult,
    *,
    region: Region | None = None,
    client: VLLMClient | None = None,
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
    retry_missing: bool = False,
) -> RegionReconciliation:
    """Reconcile an in-memory `ExtractionResult` from Stage 3."""
    return reconcile_region_data(
        result.region_id,
        result.skeleton,
        result.records,
        region=region,
        client=client,
        formats=formats,
        row_roles=row_roles,
        retry_missing=retry_missing,
    )


def assemble_sheet_reconciliation(
    sheet_name: str,
    region_results: list[RegionReconciliation],
) -> SheetReconciliation:
    """Combine per-region reconciliations into a validated sheet-level `ImportExtraction`."""
    ordered = sorted(region_results, key=lambda r: r.region_id)
    all_categories: dict[str, ImportCategory] = {}
    all_items: list[ImportItem] = []
    seen_ignored: set[tuple[int, str]] = set()
    all_ignored: list[IgnoredRow] = []

    for rr in ordered:
        for cat in rr.categories:
            existing = all_categories.get(cat.name)
            if existing is None or (existing.rate is None and cat.rate is not None):
                all_categories[cat.name] = cat
        all_items.extend(rr.items)
        for ig in rr.ignored_rows:
            key = (ig.source_row, ig.reason)
            if key not in seen_ignored:
                seen_ignored.add(key)
                all_ignored.append(ig)

    sheet_errors: list[str] = []
    extraction: ImportExtraction | None = None
    try:
        extraction = ImportExtraction(
            categories=list(all_categories.values()),
            items=all_items,
            ignored_rows=all_ignored,
        )
    except ValidationError as exc:
        for err in exc.errors():
            loc = ".".join(str(p) for p in err["loc"])
            sheet_errors.append(f"{loc}: {err['msg']}")

    return SheetReconciliation(
        sheet_name=sheet_name,
        regions=ordered,
        extraction=extraction,
        sheet_errors=sheet_errors,
    )


def _md_cell(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("|", "\\|").replace("\n", "<br>")


def reconciliation_to_markdown(reconciliation: SheetReconciliation) -> str:
    """Render a `SheetReconciliation` as a human-readable Markdown visual review."""
    total_valid = sum(len(r.items) for r in reconciliation.regions)
    total_ignored = sum(len(r.ignored_rows) for r in reconciliation.regions)
    total_issues = sum(len(r.validation_issues) for r in reconciliation.regions)
    acct_status = "PASS" if reconciliation.all_accounting_passed else "FAIL"

    lines: list[str] = [
        f"# Extraction Review — {reconciliation.sheet_name}",
        "",
        "## Summary",
        "",
        "| Metric | Value |",
        "|---|---|",
        f"| Row-Accounting Invariant | **{acct_status}** |",
        f"| Regions Reconciled | {len(reconciliation.regions)} |",
        f"| Categories | {len(reconciliation.extraction.categories) if reconciliation.extraction else 0} |",
        f"| Valid Items | {total_valid} |",
        f"| Ignored Rows | {total_ignored} |",
        f"| Validation Issues | {total_issues} |",
        "",
    ]

    if reconciliation.sheet_errors:
        lines.append("## Sheet-Level Errors")
        lines.append("")
        for err in reconciliation.sheet_errors:
            lines.append(f"- {err}")
        lines.append("")

    # Validation issues & accounting failures up front for fast review
    failed_acct = [r for r in reconciliation.regions if not r.accounting.passed]
    all_issues = [iss for r in reconciliation.regions for iss in r.validation_issues]

    if failed_acct:
        lines.extend(
            [
                "## Row-Accounting Failures",
                "",
                "| Region | Input Rows | Accounted | Missing | Hallucinated | Duplicates | Invalid | Overlap | Retried |",
                "|---|---:|---:|---|---|---|---|---|---|",
            ]
        )
        for r in failed_acct:
            a = r.accounting
            lines.append(
                f"| R{r.region_id:02d} | {a.input_row_count} | {a.accounted_row_count} | "
                f"{a.missing_rows} | {a.hallucinated_rows} | {a.duplicate_item_rows} | "
                f"{a.invalid_item_rows} | {a.overlap_rows} | {r.retried_rows} |"
            )
        lines.append("")

    if all_issues:
        lines.extend(
            [
                "## Validation Issues",
                "",
                "| Region | Row | Model | Field | Error | Raw Value |",
                "|---|---:|---|---|---|---|",
            ]
        )
        for iss in all_issues:
            lines.append(
                f"| R{iss.region_id:02d} | {iss.source_row or ''} | {_md_cell(iss.model)} | "
                f"`{_md_cell(iss.field)}` | {_md_cell(iss.error)} | `{_md_cell(iss.raw_value)}` |"
            )
        lines.append("")

    # Group valid items by category -> group
    grouped: dict[str, dict[str, list[ImportItem]]] = {}
    cat_rates: dict[str, float | None] = {}
    if reconciliation.extraction is not None:
        for cat in reconciliation.extraction.categories:
            cat_rates[cat.name] = cat.rate
            grouped.setdefault(cat.name, {})
        for item in reconciliation.extraction.items:
            grp_name = item.group or "(No Group)"
            grouped.setdefault(item.category, {}).setdefault(grp_name, []).append(item)
    else:
        for r in reconciliation.regions:
            for item in r.items:
                grp_name = item.group or "(No Group)"
                grouped.setdefault(item.category, {}).setdefault(grp_name, []).append(item)

    lines.extend(["## Extracted Items by Category", ""])
    for cat_name, groups in grouped.items():
        rate = cat_rates.get(cat_name)
        rate_str = f" *(rate: {rate})*" if rate is not None else ""
        cat_count = sum(len(items) for items in groups.values())
        lines.append(f"### {_md_cell(cat_name)}{rate_str} — {cat_count} items")
        lines.append("")

        if not groups:
            lines.append("_No valid items in this category._")
            lines.append("")
            continue

        for grp_name, items in groups.items():
            if grp_name != "(No Group)" or len(groups) > 1:
                lines.append(f"#### {_md_cell(grp_name)}")
                lines.append("")
            lines.extend(
                [
                    "| Row | ID | Name | Unit Price | Note |",
                    "|---:|---|---|---:|---|",
                ]
            )
            for it in items:
                price_str = f"{it.unit_price:,}" if it.unit_price is not None else ""
                lines.append(
                    f"| {it.source_row} | {_md_cell(it.id)} | {_md_cell(it.name)} | "
                    f"{price_str} | {_md_cell(it.note)} |"
                )
            lines.append("")

    # Compact per-region breakdown at the bottom
    lines.extend(
        [
            "## Per-Region Breakdown",
            "",
            "| Region | Accounting | Input Rows | Valid Items | Ignored | Issues |",
            "|---|---|---:|---:|---:|---:|",
        ]
    )
    for r in reconciliation.regions:
        status = "PASS" if r.accounting.passed else "FAIL"
        lines.append(
            f"| R{r.region_id:02d} | {status} | {r.accounting.input_row_count} | "
            f"{len(r.items)} | {len(r.ignored_rows)} | {len(r.validation_issues)} |"
        )
    lines.append("")

    return "\n".join(lines)


def save_reconciliation(
    reconciliation: SheetReconciliation,
    output_dir: str | Path,
    *,
    pg_table: str = "equipment_catalog",
) -> dict[str, Path]:
    """Save reconciliation report, validated Tier-2 JSON, PostgreSQL recordset JSON, and Markdown review."""
    from excel_parsing.exporters.postgres import build_pg_recordset_payload

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    report_path = out / "_reconcile_report.json"
    report_path.write_text(
        json.dumps(reconciliation.to_report_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    md_path = out / "_review.md"
    md_path.write_text(
        reconciliation_to_markdown(reconciliation),
        encoding="utf-8",
    )

    saved: dict[str, Path] = {"report": report_path, "markdown": md_path}
    if reconciliation.extraction is not None:
        import_path = out / "_validated_import.json"
        import_path.write_text(
            reconciliation.extraction.model_dump_json(indent=2),
            encoding="utf-8",
        )
        saved["validated_import"] = import_path

        pg_payload = build_pg_recordset_payload(reconciliation, table=pg_table)
        pg_path = pg_payload.save(out / "_pg_recordset.json")
        saved["pg_recordset"] = pg_path

    return saved


def reconcile_saved_extraction(extraction_dir: str | Path) -> SheetReconciliation:
    """Reconcile already-saved Stage 3 JSON files in `outputs/extraction/<sheet>/`."""
    base = Path(extraction_dir)
    summary_path = base / "_summary.json"
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    sheet_name = summary.get("sheet_name", base.name)

    region_results: list[RegionReconciliation] = []

    for entry in summary.get("regions", []):
        if entry.get("status") != "extracted":
            continue
        rid = entry["region_id"]
        skel_raw = json.loads((base / f"R{rid:02d}_skeleton.json").read_text(encoding="utf-8"))
        recs_raw = json.loads((base / f"R{rid:02d}_records.json").read_text(encoding="utf-8"))

        skeleton = [RowClassification(**r) for r in skel_raw]
        records = [ExtractedRecord(**r) for r in recs_raw]
        region_results.append(reconcile_region_data(rid, skeleton, records))

    return assemble_sheet_reconciliation(sheet_name, region_results)
