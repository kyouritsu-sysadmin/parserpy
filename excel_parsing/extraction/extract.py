# Stage 3: Hierarchical State-Machine Extraction (vLLM OpenAI-compatible endpoint)

from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass, field
import http.client
import json
from pathlib import Path
import re
import threading
import time
from typing import Any, Literal
import urllib.parse

from excel_parsing.exceptions import VLLMHTTPError, VLLMOutputTruncatedError
from excel_parsing.preprocess.detection import Region
from excel_parsing.preprocess.formatting import CellFormat, RowRole, classify_region_rows
from excel_parsing.serialize.serialize import region_to_html

SkeletonRole = Literal[
    "category_header",
    "group_header",
    "column_header",
    "line_item",
    "subtotal",
    "footnote",
    "other",
]

VALID_SKELETON_ROLES: frozenset[str] = frozenset(
    {
        "category_header",
        "group_header",
        "column_header",
        "line_item",
        "subtotal",
        "footnote",
        "other",
    }
)

_JSON_FENCE_RE = re.compile(r"```(?:json)?\s*(.*?)\s*```", re.DOTALL | re.IGNORECASE)


# ---------------------------------------------------------------------------
# vLLM OpenAI-compatible client (zero external dependencies)
# ---------------------------------------------------------------------------


_RETRYABLE_HTTP_CODES: frozenset[int] = frozenset({408, 429, 500, 502, 503, 504})


class VLLMClient:
    """Minimal persistent keep-alive client for a local vLLM OpenAI-compatible server."""

    def __init__(
        self,
        base_url: str = "http://localhost:8005/v1",
        model: str | None = None,
        api_key: str = "EMPTY",
        timeout: float = 180.0,
        default_max_tokens: int = 8192,
        disable_thinking: bool = True,
        max_retries: int = 2,
        max_concurrent_requests: int | None = 6,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        parsed = urllib.parse.urlsplit(self.base_url)
        self._scheme = parsed.scheme or "http"
        self._host = parsed.hostname or "localhost"
        self._port = parsed.port or (443 if self._scheme == "https" else 80)
        self._base_path = parsed.path.rstrip("/")
        self.model = model
        self.api_key = api_key
        self.timeout = timeout
        self.default_max_tokens = default_max_tokens
        self.disable_thinking = disable_thinking
        self.max_retries = max_retries
        self._supports_json_schema = True
        self._model_lock = threading.Lock()
        self._tls = threading.local()
        self._semaphore = (
            threading.Semaphore(max_concurrent_requests)
            if max_concurrent_requests is not None and max_concurrent_requests > 0
            else None
        )

    def _headers(self) -> dict[str, str]:
        return {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
            "Connection": "keep-alive",
        }

    def _get_conn(self, *, reset: bool = False) -> http.client.HTTPConnection:
        conn: http.client.HTTPConnection | None = getattr(self._tls, "conn", None)
        if conn is not None and reset:
            try:
                conn.close()
            except Exception:
                pass
            conn = None
        if conn is None:
            if self._scheme == "https":
                conn = http.client.HTTPSConnection(self._host, self._port, timeout=self.timeout)
            else:
                conn = http.client.HTTPConnection(self._host, self._port, timeout=self.timeout)
            self._tls.conn = conn
        return conn

    def _send_once(self, method: str, endpoint: str, data: bytes | None = None) -> dict[str, Any]:
        path = f"{self._base_path}/{endpoint.lstrip('/')}"
        headers = self._headers()
        if data is not None:
            headers["Content-Length"] = str(len(data))

        for reconnect in (False, True):
            conn = self._get_conn(reset=reconnect)
            try:
                conn.request(method, path, body=data, headers=headers)
                resp = conn.getresponse()
                raw = resp.read().decode("utf-8", errors="replace")
                if 200 <= resp.status < 300:
                    return json.loads(raw)
                raise VLLMHTTPError(resp.status, raw)
            except (
                http.client.RemoteDisconnected,
                http.client.CannotSendRequest,
                http.client.BadStatusLine,
                BrokenPipeError,
                ConnectionResetError,
            ):
                self._get_conn(reset=True)
                if not reconnect:
                    continue
                raise

        raise RuntimeError("Unreachable keep-alive retry state")

    def resolve_model(self) -> str:
        """Auto-detect the served model name from `/v1/models` if not explicitly set."""
        if self.model:
            return self.model
        with self._model_lock:
            if self.model:
                return self.model
            payload = self._send_once("GET", "/models")
            models = payload.get("data") or []
            if not models:
                raise RuntimeError(f"No models returned by {self.base_url}/models")
            self.model = models[0]["id"]
            return self.model

    def _post_chat(self, body: dict[str, Any]) -> tuple[str | None, str | None, bool]:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        last_exc: Exception | None = None

        for attempt in range(self.max_retries + 1):
            try:
                if self._semaphore is not None:
                    with self._semaphore:
                        payload = self._send_once("POST", "/chat/completions", data)
                else:
                    payload = self._send_once("POST", "/chat/completions", data)

                choice = payload["choices"][0]
                message = choice.get("message") or {}
                content = message.get("content")
                finish_reason = choice.get("finish_reason")
                has_reasoning = bool(message.get("reasoning_content") or message.get("reasoning"))
                return content, finish_reason, has_reasoning
            except VLLMHTTPError as exc:
                if exc.code in _RETRYABLE_HTTP_CODES and attempt < self.max_retries:
                    last_exc = exc
                    time.sleep(1.5 * (attempt + 1))
                    continue
                raise
            except (TimeoutError, OSError, http.client.HTTPException) as exc:
                self._get_conn(reset=True)
                last_exc = exc
                if attempt < self.max_retries:
                    time.sleep(1.5 * (attempt + 1))
                    continue
                raise

        assert last_exc is not None
        raise last_exc

    def _decode_or_raise(
        self,
        content: str | None,
        finish_reason: str | None,
        has_reasoning: bool,
        max_tokens: int,
    ) -> dict[str, Any]:
        if not content or not content.strip():
            hint = (
                " (model spent tokens in reasoning_content — try VLLMClient(disable_thinking=True) "
                "or higher default_max_tokens / smaller max_rows_per_chunk)"
                if has_reasoning
                else " (try higher default_max_tokens or smaller max_rows_per_chunk)"
            )
            if finish_reason == "length":
                raise VLLMOutputTruncatedError(
                    f"vLLM returned empty content with finish_reason='length' (max_tokens={max_tokens}){hint}"
                )
            raise ValueError(
                f"vLLM returned empty content (finish_reason={finish_reason!r}, has_reasoning={has_reasoning}){hint}"
            )

        try:
            return _parse_json_response(content)
        except json.JSONDecodeError as exc:
            if finish_reason == "length":
                raise VLLMOutputTruncatedError(
                    f"vLLM JSON output was truncated at max_tokens={max_tokens}"
                ) from exc
            raise

    def chat_json(
        self,
        system_prompt: str,
        user_prompt: str,
        *,
        json_schema: dict[str, Any] | None = None,
        schema_name: str = "response",
        temperature: float = 0.0,
        max_tokens: int | None = None,
    ) -> dict[str, Any]:
        """Call `/v1/chat/completions` and parse a JSON object response.

        Uses vLLM structured output (`json_schema`) when supported, caching fallback
        capability if an endpoint rejects schema-constrained decoding.
        """
        model_id = self.resolve_model()
        effective_max_tokens = max_tokens if max_tokens is not None else self.default_max_tokens
        base_body: dict[str, Any] = {
            "model": model_id,
            "temperature": temperature,
            "max_tokens": effective_max_tokens,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        if self.disable_thinking:
            base_body["chat_template_kwargs"] = {"enable_thinking": False}

        last_parse_err: Exception | None = None
        for parse_attempt in range(self.max_retries + 1):
            try:
                if json_schema is not None and self._supports_json_schema:
                    body = {
                        **base_body,
                        "response_format": {
                            "type": "json_schema",
                            "json_schema": {
                                "name": schema_name,
                                "strict": True,
                                "schema": json_schema,
                            },
                        },
                    }
                    try:
                        content, finish_reason, has_reasoning = self._post_chat(body)
                        return self._decode_or_raise(
                            content, finish_reason, has_reasoning, effective_max_tokens
                        )
                    except VLLMHTTPError as exc:
                        if exc.code not in (400, 422):
                            raise
                        self._supports_json_schema = False

                fallback_body = {
                    **base_body,
                    "response_format": {"type": "json_object"},
                }
                try:
                    content, finish_reason, has_reasoning = self._post_chat(fallback_body)
                except VLLMHTTPError as exc:
                    if exc.code not in (400, 422):
                        raise
                    content, finish_reason, has_reasoning = self._post_chat(base_body)

                return self._decode_or_raise(
                    content, finish_reason, has_reasoning, effective_max_tokens
                )
            except (json.JSONDecodeError, ValueError) as exc:
                if isinstance(exc, VLLMOutputTruncatedError):
                    raise
                last_parse_err = exc
                if parse_attempt < self.max_retries:
                    continue
                raise

        assert last_parse_err is not None
        raise last_parse_err


def _parse_json_response(text: str | None) -> dict[str, Any]:
    if not text:
        raise ValueError("Empty LLM response text")
    stripped = text.strip()
    fence = _JSON_FENCE_RE.search(stripped)
    if fence:
        stripped = fence.group(1).strip()

    try:
        parsed = json.loads(stripped)
    except json.JSONDecodeError:
        start = stripped.find("{")
        end = stripped.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise
        parsed = json.loads(stripped[start : end + 1])

    if not isinstance(parsed, dict):
        raise ValueError(f"Expected JSON object from LLM, got {type(parsed).__name__}")
    return parsed


# ---------------------------------------------------------------------------
# Data models for Pass 1, Pass 2, and Deterministic Tree Assembly
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RowClassification:
    row: int
    role: SkeletonRole
    label: str = ""


@dataclass(frozen=True)
class ParentState:
    category_row: int | None = None
    category_label: str | None = None
    group_row: int | None = None
    group_label: str | None = None
    column_header_rows: tuple[int, ...] = ()
    column_headers_html: str = ""

    def to_prompt_block(self) -> str:
        lines = [
            "ACTIVE PARENT CONTEXT (inherited from earlier rows):",
            f"- current_category: {self.category_label!r} (row={self.category_row})",
            f"- current_group: {self.group_label!r} (row={self.group_row})",
            f"- active_column_header_rows: {list(self.column_header_rows)}",
        ]
        if self.column_headers_html:
            lines.append("- active_column_headers_html:")
            lines.append(self.column_headers_html)
        return "\n".join(lines)


@dataclass
class StatefulChunk:
    chunk_id: int
    region_id: int
    start_row: int
    end_row: int
    rows: list[int]
    parent_state: ParentState
    html: str


@dataclass
class ExtractedRecord:
    source_row: int
    record_type: Literal["line_item", "subtotal"]
    name: str
    parent_row: int | None = None
    category: str | None = None
    group: str | None = None
    fields: dict[str, Any] = field(default_factory=dict)
    source_cells: list[str] = field(default_factory=list)


@dataclass
class GroupNode:
    group_row: int | None
    name: str
    items: list[ExtractedRecord] = field(default_factory=list)
    subtotals: list[ExtractedRecord] = field(default_factory=list)


@dataclass
class CategoryNode:
    category_row: int | None
    name: str
    groups: list[GroupNode] = field(default_factory=list)
    subtotals: list[ExtractedRecord] = field(default_factory=list)


@dataclass
class RegionTree:
    sheet_name: str
    region_id: int
    coord_range: str
    column_header_rows: list[int] = field(default_factory=list)
    categories: list[CategoryNode] = field(default_factory=list)
    footnotes: list[RowClassification] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ChunkTiming:
    chunk_id: int
    start_row: int
    end_row: int
    row_count: int
    llm_seconds: float
    records_extracted: int
    bisected: bool = False


@dataclass
class RegionTiming:
    pass1_llm_seconds: float = 0.0
    pass2_llm_seconds: float = 0.0
    assembly_seconds: float = 0.0
    total_seconds: float = 0.0
    llm_calls: int = 0
    bisects: int = 0
    chunks: list[ChunkTiming] = field(default_factory=list)


@dataclass
class ExtractionResult:
    region_id: int
    skeleton: list[RowClassification]
    chunks: list[StatefulChunk]
    records: list[ExtractedRecord]
    tree: RegionTree
    timing: RegionTiming = field(default_factory=RegionTiming)


# ---------------------------------------------------------------------------
# Pass 1: Structural Skeleton / Row Classification
# ---------------------------------------------------------------------------


_SKELETON_SYSTEM_PROMPT = """\
You are a spreadsheet structure classifier.
Classify EVERY listed row into exactly one role:
- "category_header": Top-level section or major category banner that starts a new logical section.
- "group_header": Sub-category or group label nested inside a category.
- "column_header": Table column header row (may span 1 or multiple stacked rows).
- "line_item": An individual data record / item row.
- "subtotal": A subtotal, total, or summary aggregation row.
- "footnote": A note, remark, legend, or footer annotation.
- "other": Metadata or decorative row that does not fit the above.

Return strict JSON matching the schema with one entry per input row.
"""

_SKELETON_JSON_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "rows": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "row": {"type": "integer"},
                    "role": {
                        "type": "string",
                        "enum": sorted(VALID_SKELETON_ROLES),
                    },
                },
                "required": ["row", "role"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["rows"],
    "additionalProperties": False,
}


def build_skeleton_lines(
    region: Region,
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
    *,
    max_preview_cells: int = 10,
    max_cell_chars: int = 40,
) -> tuple[list[int], str]:
    """Build a compact 1-line-per-row structural summary for Pass 1 classification."""
    fmts = formats or {}
    heuristics = row_roles if row_roles is not None else classify_region_rows(region, fmts)

    row_indices: list[int] = []
    lines: list[str] = []

    for r in range(region.min_row, region.max_row + 1):
        row_nodes = [
            region.cells[(r, c)]
            for c in range(region.min_col, region.max_col + 1)
            if (r, c) in region.cells
        ]
        if not row_nodes:
            continue

        row_indices.append(r)
        # Keep only merge origins (or non-merged cells) for concise label preview
        origin_nodes = [n for n in row_nodes if n.is_merge_origin] or row_nodes
        row_fmts = [fmts.get((n.row, n.col), CellFormat()) for n in origin_nodes]

        flags: list[str] = []
        if any(f.is_bold for f in row_fmts):
            flags.append("bold")
        if any(f.has_fill for f in row_fmts):
            flags.append("fill")
        if any(f.borders.has_bottom_emphasis for f in row_fmts):
            flags.append("bottom_border")
        max_indent = max((f.indent for f in row_fmts), default=0)
        if max_indent > 0:
            flags.append(f"indent={max_indent}")
        max_merge_w = max(
            (n.merge.col_span for n in origin_nodes if n.merge is not None),
            default=1,
        )
        if max_merge_w > 1:
            flags.append(f"merge_w={max_merge_w}")
        numeric_cnt = sum(
            1
            for n in origin_nodes
            if isinstance(n.value, (int, float)) and not isinstance(n.value, bool)
        )
        if numeric_cnt > 0:
            flags.append(f"nums={numeric_cnt}")

        preview_items = [
            str(n.value).replace("\n", " ")[:max_cell_chars]
            for n in origin_nodes[:max_preview_cells]
        ]
        preview = " | ".join(preview_items)
        flag_str = ",".join(flags) if flags else "plain"
        lines.append(
            f"row={r} | hint={heuristics.get(r, 'data')} | flags=[{flag_str}] | text=\"{preview}\""
        )

    return row_indices, "\n".join(lines)


_HEURISTIC_TO_SKELETON_FALLBACK: dict[str, SkeletonRole] = {
    "section": "category_header",
    "header": "column_header",
    "total": "subtotal",
    "data": "line_item",
}


def _is_formula_zero_row(region: Region, r: int) -> bool:
    """True when a row has no merge-origin content or only numeric `0` origin cells
    (e.g. blank template row with `SUM=0`, or second row of a 2-row category banner merge).
    """
    origin_nodes = [
        region.cells[(r, c)]
        for c in range(region.min_col, region.max_col + 1)
        if (r, c) in region.cells and region.cells[(r, c)].is_merge_origin
    ]
    if not origin_nodes:
        return True
    return all(
        isinstance(n.value, (int, float)) and not isinstance(n.value, bool) and n.value == 0
        for n in origin_nodes
    )


def _derive_row_label(region: Region, r: int) -> str:
    """Deterministically extract the primary text label of row `r` from `region.cells`."""
    origin_nodes = [
        region.cells[(r, c)]
        for c in range(region.min_col, region.max_col + 1)
        if (r, c) in region.cells and region.cells[(r, c)].is_merge_origin
    ]
    str_parts = [
        str(n.value).replace("\n", " ").strip()
        for n in origin_nodes
        if isinstance(n.value, str) and str(n.value).strip()
    ]
    if str_parts:
        return " | ".join(str_parts[:3]) if len(str_parts) > 1 else str_parts[0]
    for n in origin_nodes:
        if n.value is not None and not (isinstance(n.value, (int, float)) and n.value == 0):
            return str(n.value).strip()
    return ""


def classify_skeleton(
    client: VLLMClient,
    region: Region,
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
    *,
    max_skeleton_rows: int = 80,
) -> list[RowClassification]:
    """Pass 1: Classify every non-empty row in `region` (zero-row fast-path + batched LLM)."""
    fmts = formats or {}
    heuristics = row_roles if row_roles is not None else classify_region_rows(region, fmts)
    row_indices, skeleton_text = build_skeleton_lines(region, fmts, heuristics)
    if not row_indices:
        return []

    line_by_row = dict(zip(row_indices, skeleton_text.splitlines(), strict=False))
    by_row: dict[int, RowClassification] = {}

    llm_rows: list[int] = []
    for r in row_indices:
        if _is_formula_zero_row(region, r):
            by_row[r] = RowClassification(row=r, role="other", label="")
        else:
            llm_rows.append(r)

    def _run_batch(batch_rows: list[int]) -> None:
        if not batch_rows:
            return
        batch_text = "\n".join(line_by_row[r] for r in batch_rows)
        user_prompt = (
            f"Sheet: {region.sheet_name} | Region #{region.region_id} ({region.coord_range})\n"
            f"Rows in this batch: {len(batch_rows)}\n\n"
            f"{batch_text}"
        )
        dyn_max_tokens = min(
            getattr(client, "default_max_tokens", 8192),
            max(384, 32 * len(batch_rows) + 128),
        )
        data = client.chat_json(
            _SKELETON_SYSTEM_PROMPT,
            user_prompt,
            json_schema=_SKELETON_JSON_SCHEMA,
            schema_name="row_skeleton",
            max_tokens=dyn_max_tokens,
        )
        for entry in data.get("rows", []):
            r = entry.get("row")
            role = entry.get("role", "other")
            if isinstance(r, int) and role in VALID_SKELETON_ROLES:
                raw_lbl = str(entry.get("label", "")).strip()
                lbl = raw_lbl or _derive_row_label(region, r)
                by_row[r] = RowClassification(row=r, role=role, label=lbl)  # type: ignore[arg-type]

    for offset in range(0, len(llm_rows), max_skeleton_rows):
        batch = llm_rows[offset : offset + max_skeleton_rows]
        _run_batch(batch)
        missing = [r for r in batch if r not in by_row]
        if missing:
            _run_batch(missing)

    # Fallback preserves heuristic row role (e.g. 'data' -> 'line_item') instead of dropping items
    result: list[RowClassification] = []
    for r in row_indices:
        if r in by_row:
            result.append(by_row[r])
        else:
            fallback_role = _HEURISTIC_TO_SKELETON_FALLBACK.get(heuristics.get(r, "data"), "line_item")
            result.append(RowClassification(row=r, role=fallback_role, label=_derive_row_label(region, r)))
    return result


# ---------------------------------------------------------------------------
# Pass 2: Stateful Chunk Slicing & Flat Record Extraction
# ---------------------------------------------------------------------------


def _subregion_for_rows(region: Region, rows: list[int]) -> Region:
    row_set = set(rows)
    sub_cells = {
        coord: node for coord, node in region.cells.items() if coord[0] in row_set
    }
    if sub_cells:
        cols = [c for _, c in sub_cells]
        min_c, max_c = min(cols), max(cols)
    else:
        min_c, max_c = region.min_col, region.max_col
    return Region(
        region_id=region.region_id,
        sheet_name=region.sheet_name,
        min_row=min(rows),
        min_col=min_c,
        max_row=max(rows),
        max_col=max_c,
        cells=sub_cells,
    )


def advance_parent_state(
    initial_state: ParentState,
    rows: list[int],
    skeleton_by_row: dict[int, RowClassification],
    region: Region,
    formats: dict[tuple[int, int], CellFormat] | None = None,
    *,
    include_coords: bool = True,
) -> ParentState:
    """Deterministically advance `ParentState` across `rows`."""
    fmts = formats or {}
    cat_row, cat_label = initial_state.category_row, initial_state.category_label
    grp_row, grp_label = initial_state.group_row, initial_state.group_label
    col_rows = list(initial_state.column_header_rows)
    col_changed = False
    prev_was_col = False

    for r in rows:
        rc = skeleton_by_row.get(r)
        if rc is None:
            continue
        if rc.role == "category_header":
            cat_row, cat_label = rc.row, rc.label
            grp_row, grp_label = None, None
            prev_was_col = False
        elif rc.role == "column_header":
            if not prev_was_col:
                col_rows = [rc.row]
            else:
                col_rows.append(rc.row)
            col_changed = True
            prev_was_col = True
        else:
            prev_was_col = False
            if rc.role == "group_header":
                grp_row, grp_label = rc.row, rc.label

    if col_changed and col_rows:
        col_sub = _subregion_for_rows(region, col_rows)
        col_html = region_to_html(
            col_sub,
            formats=fmts,
            row_roles={r: "header" for r in col_rows},
            rows=col_rows,
            include_coords=include_coords,
        )
    else:
        col_html = initial_state.column_headers_html

    return ParentState(
        category_row=cat_row,
        category_label=cat_label,
        group_row=grp_row,
        group_label=grp_label,
        column_header_rows=tuple(col_rows),
        column_headers_html=col_html,
    )


def slice_stateful_chunks(
    region: Region,
    skeleton: list[RowClassification],
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
    *,
    max_rows_per_chunk: int = 25,
    include_coords: bool = True,
) -> list[StatefulChunk]:
    """Slice `region` strictly at `category_header` boundaries (and `max_rows_per_chunk`),
    capturing the active `ParentState` at the start of every chunk.
    """
    if not skeleton:
        return []

    fmts = formats or {}
    chunks: list[StatefulChunk] = []

    current_cat: tuple[int, str] | None = None
    current_grp: tuple[int, str] | None = None
    active_col_rows: list[int] = []
    active_col_html = ""

    buf_rows: list[int] = []
    buf_state = ParentState()

    def _flush() -> None:
        nonlocal buf_rows
        if not buf_rows:
            return
        sub = _subregion_for_rows(region, buf_rows)
        chunk_html = region_to_html(
            sub,
            formats=fmts,
            row_roles=row_roles,
            rows=buf_rows,
            include_coords=include_coords,
        )
        chunks.append(
            StatefulChunk(
                chunk_id=len(chunks) + 1,
                region_id=region.region_id,
                start_row=buf_rows[0],
                end_row=buf_rows[-1],
                rows=list(buf_rows),
                parent_state=buf_state,
                html=chunk_html,
            )
        )
        buf_rows = []

    def _snapshot_state() -> ParentState:
        return ParentState(
            category_row=current_cat[0] if current_cat else None,
            category_label=current_cat[1] if current_cat else None,
            group_row=current_grp[0] if current_grp else None,
            group_label=current_grp[1] if current_grp else None,
            column_header_rows=tuple(active_col_rows),
            column_headers_html=active_col_html,
        )

    prev_was_col_header = False

    for rc in skeleton:
        # Start a new chunk strictly on category_header boundaries
        if rc.role == "category_header":
            _flush()
            current_cat = (rc.row, rc.label)
            current_grp = None
            buf_state = _snapshot_state()
            buf_rows.append(rc.row)
            prev_was_col_header = False
            continue

        # Update active column headers (replace if a new non-contiguous header block starts)
        if rc.role == "column_header":
            if not prev_was_col_header:
                active_col_rows = [rc.row]
            else:
                active_col_rows.append(rc.row)
            col_sub = _subregion_for_rows(region, active_col_rows)
            active_col_html = region_to_html(
                col_sub,
                formats=fmts,
                row_roles={r: "header" for r in active_col_rows},
                rows=active_col_rows,
                include_coords=include_coords,
            )
            if not buf_rows:
                buf_state = _snapshot_state()
            buf_rows.append(rc.row)
            prev_was_col_header = True
            continue

        prev_was_col_header = False

        # Size-based split inside a very large category (snapshot BEFORE mutating current_grp)
        if len(buf_rows) >= max_rows_per_chunk and rc.role in ("group_header", "line_item"):
            _flush()
            buf_state = _snapshot_state()

        if not buf_rows:
            buf_state = _snapshot_state()

        if rc.role == "group_header":
            current_grp = (rc.row, rc.label)

        buf_rows.append(rc.row)

    _flush()
    return chunks


_CHUNK_SYSTEM_PROMPT = """\
You extract structured records from a compact HTML slice of an Excel sheet.
Rules:
1. Emit ONLY flat records for `line_item` rows listed in ROW ROLES IN THIS CHUNK. Do NOT emit headers, subtotals, or nested trees.
2. Tag every record with its exact integer `source_row` (from `c="B15"` cell coordinates in the HTML).
3. Tag `parent_row` with the nearest governing `group_header` row (or `category_header` row if no group applies; null if none).
4. Use ACTIVE PARENT CONTEXT when a category, group, or column header was defined in an earlier chunk.
5. Put meaningful non-empty attributes into `fields` keyed by active column header names (e.g. No/id, 単価/原価/unit_price, 容量, 電圧, 備考/note). Omit repeated `0` quantity/amount columns (such as `SS-1..SS-20` or `金額=0`).
6. If a table has NO `column_header` row and a `line_item` row has `[item_name, numeric_price]`, store the numeric price under `fields.unit_price`. Preserve numbers as numbers and keep `[ERROR:...]` tags verbatim.
Return strict JSON matching the schema.
"""

_CHUNK_JSON_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "records": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "source_row": {"type": "integer"},
                    "name": {"type": "string"},
                    "parent_row": {"type": ["integer", "null"]},
                    "category": {"type": ["string", "null"]},
                    "group": {"type": ["string", "null"]},
                    "fields": {"type": "object"},
                },
                "required": [
                    "source_row",
                    "name",
                    "parent_row",
                    "category",
                    "group",
                    "fields",
                ],
                "additionalProperties": False,
            },
        }
    },
    "required": ["records"],
    "additionalProperties": False,
}


def _row_source_cells(region: Region | None, source_row: int) -> list[str]:
    if region is None:
        return []
    coords: list[str] = []
    for c in range(region.min_col, region.max_col + 1):
        node = region.cells.get((source_row, c))
        if node is not None and node.is_merge_origin:
            coords.append(node.coord)
    return coords


def extract_chunk_records(
    client: VLLMClient,
    chunk: StatefulChunk,
    skeleton_by_row: dict[int, RowClassification],
    *,
    region: Region | None = None,
) -> list[ExtractedRecord]:
    """Pass 2: Extract flat `line_item` records from a single stateful HTML chunk."""
    chunk_roles = "\n".join(
        f"  row={r}: {skeleton_by_row[r].role} ({skeleton_by_row[r].label!r})"
        for r in chunk.rows
        if r in skeleton_by_row
    )
    extractable_cnt = sum(
        1
        for r in chunk.rows
        if r in skeleton_by_row and skeleton_by_row[r].role == "line_item"
    )
    dyn_max_tokens = min(
        getattr(client, "default_max_tokens", 8192),
        max(768, 160 * max(1, extractable_cnt) + 256),
    )

    user_prompt = (
        f"Chunk #{chunk.chunk_id} (rows {chunk.start_row}..{chunk.end_row})\n\n"
        f"{chunk.parent_state.to_prompt_block()}\n\n"
        f"ROW ROLES IN THIS CHUNK:\n{chunk_roles}\n\n"
        f"CHUNK HTML:\n{chunk.html}"
    )

    data = client.chat_json(
        _CHUNK_SYSTEM_PROMPT,
        user_prompt,
        json_schema=_CHUNK_JSON_SCHEMA,
        schema_name="chunk_records",
        max_tokens=dyn_max_tokens,
    )

    records: list[ExtractedRecord] = []
    for item in data.get("records", []):
        src_row = item.get("source_row")
        if not isinstance(src_row, int):
            continue
        rec_type = item.get("record_type", "line_item")
        if rec_type not in ("line_item", "subtotal"):
            rec_type = "line_item"

        src_cells = (
            [str(c) for c in item["source_cells"]]
            if isinstance(item.get("source_cells"), list) and item["source_cells"]
            else _row_source_cells(region, src_row)
        )
        records.append(
            ExtractedRecord(
                source_row=src_row,
                record_type=rec_type,  # type: ignore[arg-type]
                name=str(item.get("name", "")).strip(),
                parent_row=item.get("parent_row") if isinstance(item.get("parent_row"), int) else None,
                category=item.get("category") or chunk.parent_state.category_label,
                group=item.get("group") or chunk.parent_state.group_label,
                fields=item.get("fields") if isinstance(item.get("fields"), dict) else {},
                source_cells=src_cells,
            )
        )
    return records


# ---------------------------------------------------------------------------
# Deterministic Tree Assembly in Python
# ---------------------------------------------------------------------------


def assemble_hierarchy(
    region: Region,
    skeleton: list[RowClassification],
    records: list[ExtractedRecord],
) -> RegionTree:
    """Deterministically build and validate a hierarchical `RegionTree` from
    Pass 1 row classifications and Pass 2 flat records.
    """
    tree = RegionTree(
        sheet_name=region.sheet_name,
        region_id=region.region_id,
        coord_range=region.coord_range,
        column_header_rows=[rc.row for rc in skeleton if rc.role == "column_header"],
        footnotes=[rc for rc in skeleton if rc.role == "footnote"],
    )

    # Index records by source_row
    records_by_row: dict[int, list[ExtractedRecord]] = {}
    for rec in records:
        records_by_row.setdefault(rec.source_row, []).append(rec)

    valid_rows = {rc.row for rc in skeleton}
    for rec in records:
        if rec.source_row not in valid_rows:
            tree.warnings.append(
                f"Record at source_row={rec.source_row} ({rec.name!r}) is not in region skeleton."
            )

    current_cat: CategoryNode | None = None
    current_grp: GroupNode | None = None

    def _ensure_category(row: int | None = None, name: str = "Uncategorized") -> CategoryNode:
        nonlocal current_cat
        if current_cat is None:
            current_cat = CategoryNode(category_row=row, name=name)
            tree.categories.append(current_cat)
        return current_cat

    def _ensure_group(row: int | None = None, name: str = "Default") -> GroupNode:
        nonlocal current_grp
        cat = _ensure_category()
        if current_grp is None:
            current_grp = GroupNode(group_row=row, name=name)
            cat.groups.append(current_grp)
        return current_grp

    for rc in skeleton:
        if rc.role == "category_header":
            current_cat = CategoryNode(category_row=rc.row, name=rc.label or f"Row {rc.row}")
            tree.categories.append(current_cat)
            current_grp = None
            continue

        if rc.role == "group_header":
            cat = _ensure_category()
            current_grp = GroupNode(group_row=rc.row, name=rc.label or f"Row {rc.row}")
            cat.groups.append(current_grp)
            continue

        if rc.role == "line_item":
            row_recs = records_by_row.get(rc.row, [])
            if not row_recs:
                tree.warnings.append(
                    f"Missing extraction for line_item row={rc.row} ({rc.label!r})."
                )
                continue
            grp = _ensure_group()
            for rec in row_recs:
                if rec.parent_row is None:
                    rec.parent_row = grp.group_row if grp.group_row is not None else _ensure_category().category_row
                if not rec.category:
                    rec.category = _ensure_category().name
                if not rec.group:
                    rec.group = grp.name
                grp.items.append(rec)
            continue

        if rc.role == "subtotal":
            cat = _ensure_category()
            row_recs = records_by_row.get(rc.row) or [
                ExtractedRecord(
                    source_row=rc.row,
                    record_type="subtotal",
                    name=rc.label or f"Subtotal {rc.row}",
                    parent_row=current_grp.group_row if current_grp is not None else cat.category_row,
                    category=cat.name,
                    group=current_grp.name if current_grp is not None else None,
                    fields={},
                    source_cells=_row_source_cells(region, rc.row),
                )
            ]
            for rec in row_recs:
                if current_grp is not None:
                    current_grp.subtotals.append(rec)
                else:
                    cat.subtotals.append(rec)

    return tree


# ---------------------------------------------------------------------------
# End-to-end Region Orchestrator
# ---------------------------------------------------------------------------


def extract_region(
    client: VLLMClient,
    region: Region,
    formats: dict[tuple[int, int], CellFormat] | None = None,
    row_roles: dict[int, RowRole] | None = None,
    *,
    max_rows_per_chunk: int = 25,
    include_coords: bool = True,
    max_chunk_workers: int = 3,
) -> ExtractionResult:
    """Run Pass 1 (skeleton) -> Pass 2 (parallel stateful chunk extraction) -> Tree Assembly.

    If a chunk hits `max_tokens` (`VLLMOutputTruncatedError`), it is automatically
    bisected into two smaller chunks with `ParentState` accurately advanced across the split.
    """
    t_total_start = time.perf_counter()

    t_p1 = time.perf_counter()
    skeleton = classify_skeleton(client, region, formats, row_roles)
    pass1_seconds = round(time.perf_counter() - t_p1, 3)
    skeleton_by_row = {rc.row: rc for rc in skeleton}

    chunks = slice_stateful_chunks(
        region,
        skeleton,
        formats=formats,
        row_roles=row_roles,
        max_rows_per_chunk=max_rows_per_chunk,
        include_coords=include_coords,
    )

    timing = RegionTiming(
        pass1_llm_seconds=pass1_seconds,
        llm_calls=1 if skeleton else 0,
    )

    def _extract_with_bisect(
        chunk: StatefulChunk,
        *,
        bisected: bool = False,
    ) -> tuple[list[ExtractedRecord], list[ChunkTiming], int, int]:
        has_extractable = any(
            skeleton_by_row[r].role == "line_item"
            for r in chunk.rows
            if r in skeleton_by_row
        )
        if not has_extractable:
            return [], [], 0, 0

        t_chunk = time.perf_counter()
        try:
            recs = extract_chunk_records(client, chunk, skeleton_by_row, region=region)
            elapsed = round(time.perf_counter() - t_chunk, 3)
            ct = ChunkTiming(
                chunk_id=chunk.chunk_id,
                start_row=chunk.start_row,
                end_row=chunk.end_row,
                row_count=len(chunk.rows),
                llm_seconds=elapsed,
                records_extracted=len(recs),
                bisected=bisected,
            )
            return recs, [ct], 1, 0
        except VLLMOutputTruncatedError:
            if len(chunk.rows) <= 1:
                raise
            mid = len(chunk.rows) // 2
            left_rows = chunk.rows[:mid]
            right_rows = chunk.rows[mid:]

            left_sub = _subregion_for_rows(region, left_rows)
            right_sub = _subregion_for_rows(region, right_rows)
            right_state = advance_parent_state(
                chunk.parent_state,
                left_rows,
                skeleton_by_row,
                region,
                formats=formats,
                include_coords=include_coords,
            )

            left_chunk = StatefulChunk(
                chunk_id=chunk.chunk_id,
                region_id=chunk.region_id,
                start_row=left_rows[0],
                end_row=left_rows[-1],
                rows=left_rows,
                parent_state=chunk.parent_state,
                html=region_to_html(
                    left_sub,
                    formats=formats,
                    row_roles=row_roles,
                    rows=left_rows,
                    include_coords=include_coords,
                ),
            )
            right_chunk = StatefulChunk(
                chunk_id=chunk.chunk_id,
                region_id=chunk.region_id,
                start_row=right_rows[0],
                end_row=right_rows[-1],
                rows=right_rows,
                parent_state=right_state,
                html=region_to_html(
                    right_sub,
                    formats=formats,
                    row_roles=row_roles,
                    rows=right_rows,
                    include_coords=include_coords,
                ),
            )
            if max_chunk_workers > 1:
                with ThreadPoolExecutor(max_workers=2) as sub_pool:
                    f_left = sub_pool.submit(_extract_with_bisect, left_chunk, bisected=True)
                    f_right = sub_pool.submit(_extract_with_bisect, right_chunk, bisected=True)
                    l_recs, l_timings, l_calls, l_bis = f_left.result()
                    r_recs, r_timings, r_calls, r_bis = f_right.result()
            else:
                l_recs, l_timings, l_calls, l_bis = _extract_with_bisect(left_chunk, bisected=True)
                r_recs, r_timings, r_calls, r_bis = _extract_with_bisect(right_chunk, bisected=True)
            return (
                l_recs + r_recs,
                l_timings + r_timings,
                1 + l_calls + r_calls,
                1 + l_bis + r_bis,
            )

    t_p2 = time.perf_counter()
    all_records: list[ExtractedRecord] = []
    workers = max(1, min(max_chunk_workers, len(chunks))) if chunks else 1

    if workers == 1:
        chunk_outputs = [_extract_with_bisect(ch) for ch in chunks]
    else:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            chunk_outputs = list(pool.map(_extract_with_bisect, chunks))

    for recs, c_timings, calls, bis in chunk_outputs:
        all_records.extend(recs)
        timing.chunks.extend(c_timings)
        timing.llm_calls += calls
        timing.bisects += bis

    timing.pass2_llm_seconds = round(time.perf_counter() - t_p2, 3)

    t_asm = time.perf_counter()
    tree = assemble_hierarchy(region, skeleton, all_records)
    timing.assembly_seconds = round(time.perf_counter() - t_asm, 4)
    timing.total_seconds = round(time.perf_counter() - t_total_start, 3)

    return ExtractionResult(
        region_id=region.region_id,
        skeleton=skeleton,
        chunks=chunks,
        records=all_records,
        tree=tree,
        timing=timing,
    )


def save_extraction_result(
    result: ExtractionResult,
    output_dir: str | Path,
) -> dict[str, Path]:
    """Persist Pass 1 skeleton, Pass 2 flat records, and assembled tree locally for review."""
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    prefix = f"R{result.region_id:02d}"

    skeleton_path = out / f"{prefix}_skeleton.json"
    records_path = out / f"{prefix}_records.json"
    tree_path = out / f"{prefix}_tree.json"

    skeleton_path.write_text(
        json.dumps([asdict(rc) for rc in result.skeleton], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    records_path.write_text(
        json.dumps([asdict(rec) for rec in result.records], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    tree_path.write_text(
        json.dumps(result.tree.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return {
        "skeleton": skeleton_path,
        "records": records_path,
        "tree": tree_path,
    }
