# Excel Parsing & Hierarchical LLM Extraction Engine

Deterministic spatial decomposition, compact semantic HTML serialization, two-pass stateful vLLM extraction, and strict row-accounting reconciliation for complex enterprise spreadsheets (multi-table sheets, merged cells, hidden template ranges, and hierarchical categories/groups).

---

## Current Pipeline Stages

1. **Stage 1 — Spatial Decomposition & Preprocessing (`preprocess/`)**
   - **Merge resolution (`decompose.py`):** Resolves merged ranges non-destructively while filtering hidden rows and hidden `ColumnDimension` ranges (`min..max`).
   - **Normalization (`normalize.py`):** Applies Unicode NFKC normalization, full-width space cleanup, ISO-8601 date formatting, and explicit Excel formula error tagging (`[ERROR:#DIV/0!]`, `[ERROR:#VALUE!]`).
   - **Connected-component region detection (`detection.py`):** Detects tables in visible coordinates, detaches sheet-wide title banners that bridge disjoint tables, splits side-by-side repeated-header tables (`No | 機器名 | No | 機器名`), and reunites columns split across blank user-input `数量` columns.
   - **Formatting & row heuristics (`formatting.py`):** Extracts bold/fill/border/indent/number-format signals and lightweight row-role hints (`section`, `header`, `total`, `data`).
2. **Stage 2 — Compact Semantic HTML Serialization (`serialize/`)**
   - Renders each detected region into minimal HTML (`<thead>`/`<tbody>`, native `rowspan`/`colspan` over visible axes, `c="A1"` cell coordinates) with empty/hidden columns stripped.
3. **Stage 3 — Two-Pass Stateful vLLM Extraction (`extraction/`)**
   - **Pass 1 (Skeleton):** Classifies every non-zero row into `category_header`, `group_header`, `column_header`, `line_item`, `subtotal`, `footnote`, or `other`.
   - **Pass 2 (Stateful Chunks):** Slices tables on category boundaries with active `ParentState` inheritance, extracts flat `line_item` records in parallel over persistent HTTP keep-alive connections, and auto-bisects on `max_tokens` truncation with accurate state propagation.
4. **Stage 4 — Reconciliation & Tier-2 Schema Validation (`reconcile/`)**
   - Enforces the **Row-Accounting Invariant** ($R_{\text{input}} = R_{\text{items}} \cup R_{\text{ignored}}$ with zero missing, hallucinated, duplicate, overlapping, or invalid rows), performs targeted single-chunk retries for missing rows, parses category rates, and validates every record against strict Pydantic domain models (`ImportExtraction`).

---

## Enterprise Wrapper Architecture (Layer 1 & Layer 2)

Because the pipeline has two distinct runtime profiles—**Stages 1–2 are CPU-only and finish in < 1s**, while **Stages 3–4 are GPU-bound against vLLM and take 60–180s**—enterprise integration separates **the stateless core engine (Layer 1)** from **the async service transport (Layer 2)** using a Ports & Adapters design.

```text
+-------------------------------------------------------------------+
| Consuming Applications / Backends / Workers / CLI / Batch Scripts |
+---------------------------------+---------------------------------+
                                  |
          In-Process Python Import|           HTTP / SSE / Webhook
                                  v                       v
+-----------------------------------+   +---------------------------+
| Layer 1: Python SDK & Exporters   |   | Layer 2: Async FastAPI    |
| - inspect_workbook()  (< 1s)      |<--| - POST /v1/inspect (sync) |
| - extract_workbook()              |   | - POST /v1/jobs    (202)  |
| - to_pg_recordset() ($1::json)    |   | - GET  /v1/jobs/{id}      |
+-----------------+-----------------+   | - Global vLLM Queue Gate  |
                  |                     +---------------------------+
                  v
+-------------------------------------------------------------------+
| Core Pipeline: preprocess -> serialize -> extraction -> reconcile |
+-------------------------------------------------------------------+
```

---

### Layer 1 Definition: Stateless Python SDK & PostgreSQL Exporter

Layer 1 packages the pipeline as a pure, installable Python library (`excel_parsing`) with **zero web-framework or database-driver lock-in** and no mandatory local filesystem side effects.

#### Core Capabilities
- **In-memory & stream inputs:** Accepts `bytes | BinaryIO | Path` directly so callers can pass uploaded files or S3/MinIO streams without temp-file boilerplate.
- **Two clean entrypoints:**
  - `inspect_workbook(source, ...) -> WorkbookInspection`: Runs Stages 1–2 only (< 1s) and returns visible sheet names, detected regions, error counts, and preview HTML.
  - `extract_workbook(source, ...) -> WorkbookExtractionResult`: Runs Stages 1–4 and returns typed per-sheet `SheetReconciliation` + validated `ImportExtraction` models + timing telemetry.
- **Parameterized PostgreSQL `json_to_recordset` Exporter (`exporters/postgres.py`):**
  Transforms a validated `ImportExtraction` into a **single denormalized items recordset** (merging each category's `rate` onto its item rows) paired with a parameterized `INSERT INTO ... SELECT ... FROM json_to_recordset($1::json)` query body.

#### Installation (Editable / Local Dependency)
```bash
# Install into the current environment (with CLI `excel-parse`)
uv pip install -e /mnt/workspace/projects/excel_parsing

# Or add as a dependency from another uv project:
uv add --editable /mnt/workspace/projects/excel_parsing
```

#### Example SDK Usage
```python
from excel_parsing import ExcelParser, ParserConfig, extract_workbook, inspect_workbook

# 1. Fast Stage 1-2 inspection (< 1s, accepts Path | str | bytes | BinaryIO)
inspection = inspect_workbook(file_bytes)

# 2. Reusable client for backend workers (shares HTTP keep-alive pool + vLLM semaphore)
parser = ExcelParser(
    ParserConfig(
        vllm_base_url="http://localhost:8005/v1",
        max_concurrent_requests=6,
    )
)
result = parser.extract(
    file_bytes,
    sheet_names=None,          # defaults to all visible sheets
    fail_on_accounting=True,   # raises RowAccountingError before DB write if invariant fails
)

# 3. Export single-sheet or multi-sheet workbook as parameterized PostgreSQL json_to_recordset
pg_payload = result.to_pg_recordset(
    table="equipment_catalog",
    column_map={"id": "catalog_id", "group": "item_group"},
    extra_columns={"import_batch_id": "$2::uuid"},
    on_conflict="(catalog_id) DO UPDATE SET unit_price = EXCLUDED.unit_price",
)
# Caller executes inside its own asyncpg / psycopg / SQLAlchemy transaction:
# await conn.execute(pg_payload.query, *pg_payload.params, batch_id)
```

#### CLI Usage (`excel-parse`)
```bash
# Fast CPU-only workbook inspection (< 1s)
uv run excel-parse inspect samples/catalogs/file.xlsx

# Full extraction -> emits parameterized json_to_recordset JSON bundle to stdout
uv run excel-parse extract samples/catalogs/file.xlsx --pg-table equipment_catalog -o outputs/
```

#### Parameterized `json_to_recordset` Output Contract
```json
{
  "query": "INSERT INTO equipment_catalog (import_batch_id, sheet_name, source_row, catalog_id, name, category, category_rate, item_group, unit_price, note) SELECT $2::uuid, $3::text, r.source_row, r.id, r.name, r.category, r.category_rate, r.\"group\", r.unit_price, r.note FROM json_to_recordset($1::json) AS r(source_row integer, id text, name text, category text, category_rate numeric(5,2), \"group\" text, unit_price integer, note text)",
  "params": [
    "[{\"source_row\":10,\"id\":null,\"name\":\"DS 400A\",\"category\":\"【DS】\",\"category_rate\":null,\"group\":null,\"unit_price\":43910,\"note\":null}]"
  ],
  "row_count": 397,
  "all_accounting_passed": true
}
```

#### Why Keep Database Mapping at the Boundary?
1. **Zero DB credentials in the parser:** The extraction engine never needs write credentials to consuming applications' databases.
2. **Caller owns the transaction:** Host applications execute the query inside their own transaction boundary, attach `tenant_id` / `import_batch_id` columns, and roll back atomically if `all_accounting_passed` is `false`.
3. **Single round-trip bulk insert & SQL-injection safety:** `json_to_recordset($1::json)` binds the entire validated array as `$1` and inserts hundreds of rows in milliseconds.

---

### Layer 2 Definition: Thin Async HTTP Microservice (FastAPI + Job Queue)

Layer 2 wraps Layer 1 for polyglot services, web frontends, and multi-application deployments. Because LLM extraction takes 60–180s per workbook, **blocking HTTP `POST /extract` calls time out behind load balancers, reverse proxies, and browsers**. Layer 2 exposes an asynchronous job API with centralized vLLM backpressure.

#### API Endpoints

| Endpoint | Mode | Description |
|---|---|---|
| `POST /v1/inspect` | **Sync (< 1s)** | Runs Stages 1–2 only. Returns visible sheets, detected regions, formula error summary, and HTML preview so UIs can confirm target sheets before spending GPU tokens. |
| `POST /v1/jobs` | **Async (`202 Accepted`)** | Accepts `.xlsx` upload (or object-storage URI), sheet filter, optional PostgreSQL target table/column mapping, and `webhook_url`. Computes SHA-256 `(file_hash + config)` for **idempotency & result caching**. |
| `GET /v1/jobs/{job_id}` | **Sync Poll** | Returns job state (`queued`, `running`, `succeeded`, `failed`), region progress (`8/14 regions`), and final `{ validated_import, reconcile_report, pg_recordset }` bundle. |
| `GET /v1/jobs/{job_id}/events` | **SSE Stream** | Streams live per-region completion and accounting status events to UIs. |
| `GET /healthz` & `/readyz` | **Sync** | Verifies service liveness, vLLM `/v1/models` reachability, and worker queue depth. |

#### Operational Guarantees in Layer 2
- **Centralized vLLM Concurrency Gate:** A single bounded worker queue (`arq` + Redis for multi-replica deployments, or an in-process bounded `asyncio.Semaphore` worker for single-node deployments) owns the global vLLM concurrency budget so simultaneous uploads from multiple apps queue cleanly instead of exhausting GPU KV-cache.
- **Content-Addressed Idempotency:** Re-submitting an identical workbook and configuration returns the cached validated payload immediately without re-running vLLM.
- **Artifact Persistence (Optional):** Saves compact HTML previews, per-region audit JSONs, and `_review.md` diagnostics to local storage or S3/MinIO keyed by `job_id`.

---

## Repository Layout

```text
excel_parsing/                  # Installable package root (pyproject.toml)
├── excel_parsing/              # Layer 1 Python SDK module (`from excel_parsing import ...`)
│   ├── __init__.py             # Public SDK exports
│   ├── exceptions.py           # ExcelParsingError, RowAccountingError, VLLMHTTPError, ...
│   ├── models.py               # ParserConfig, WorkbookInspection, WorkbookExtractionResult, ...
│   ├── pipeline.py             # ExcelParser, inspect_workbook(), extract_workbook()
│   ├── cli.py                  # `excel-parse` CLI (`inspect` & `extract`)
│   ├── exporters/
│   │   └── postgres.py         # Parameterized json_to_recordset($1::json) payload builder
│   ├── preprocess/             # Stage 1: decompose, normalize, detection, formatting
│   ├── serialize/              # Stage 2: compact semantic HTML
│   ├── extraction/             # Stage 3: VLLMClient + Pass 1 / Pass 2 state machine
│   └── reconcile/              # Stage 4: Row-accounting invariant + Tier-2 Pydantic schemas
├── tests/
│   └── test_sdk.py             # SDK unit & integration tests
└── main.py                     # Local example runner
```

---

## What Was Built (Layer 1 SDK — v0.2.0)

- **Installable SDK Package (`pyproject.toml`, `excel_parsing/__init__.py`):**
  Packaged all stages inside the `excel_parsing` namespace (`hatchling` build backend) so there are zero top-level module collisions when installed into external projects via `uv add --editable /path/to/excel_parsing` or `uv pip install -e .`.
- **Stateless Engine & Reusable Client (`excel_parsing/pipeline.py`, `models.py`, `exceptions.py`):**
  - Accepts `Path | str | bytes | bytearray | BinaryIO` directly in memory without temporary files.
  - `ExcelParser.inspect()` / `inspect_workbook()`: CPU-only Stages 1–2 inspection in ~40–150 ms.
  - `ExcelParser.extract()` / `extract_workbook()`: Full Stages 1–4 extraction sharing a single `VLLMClient` HTTP keep-alive connection pool and concurrency semaphore, with `fail_on_accounting=True` (`RowAccountingError`), `content_sha256` fingerprinting, and `on_progress` callbacks.
- **PostgreSQL `json_to_recordset` Exporter (`excel_parsing/exporters/postgres.py`):**
  - Denormalizes `ImportExtraction` into a single flat items array with `category_rate` merged onto each item row.
  - Emits `PostgresRecordsetPayload` (`query` + `params: [json_array_str]`) supporting `asyncpg` (`$1::json`), `psycopg` (`%s::json`), and SQLAlchemy (`:records`), plus `column_map`, `extra_columns`, `include_fields`, `on_conflict`, `returning`, `jsonb_to_recordset`, and safe quoting of SQL reserved keywords (such as `"group"`).
  - Automatically persists `_pg_recordset.json` per sheet and `_workbook_pg_recordset.json` per workbook when `output_dir` is provided.
- **Core Pipeline & vLLM Throughput Optimizations (`preprocess/`, `serialize/`, `extraction/`, `reconcile/`):**
  - Strips hidden rows and openpyxl `ColumnDimension` hidden ranges (`108` hidden columns removed in multi-island templates, reducing HTML size by >60%).
  - Splits banner-bridged and side-by-side multi-up tables (`No | 機器名 | No | 機器名`) and reunites columns split across blank `数量` columns.
  - Uses thread-local HTTP keep-alive sockets, dynamic per-chunk `max_tokens`, slimmed Pass 1/Pass 2 schemas, stateful bisection (`advance_parent_state`), category-rate parsing, and strict row-accounting validation (`100%` pass rate across `初回` and `その他機器`).
- **CLI & Test Suite (`excel_parsing/cli.py`, `tests/test_sdk.py`, `main.py`):**
  - `excel-parse inspect <file.xlsx>` and `excel-parse extract <file.xlsx> --pg-table ...` CLI commands.
  - Unit and offline reconciliation test suite (`python -m unittest discover -s tests -v`).

---

## Implementation Status & Roadmap

### Phase 1 — Layer 1: Core SDK & PostgreSQL `json_to_recordset` Exporter (Completed)
1. **Built `excel_parsing/exporters/postgres.py`:**
   - Denormalizes `ImportExtraction` (`items` joined with `categories[name].rate` -> `category_rate`).
   - Generates parameterized `INSERT INTO <table> (...) SELECT ... FROM json_to_recordset($1::json) AS r(...)` SQL + JSON parameter array with custom table/column mapping, extra bind columns, and `ON CONFLICT` / `RETURNING` support.
   - Persists `_pg_recordset.json` and `_workbook_pg_recordset.json` automatically alongside `_validated_import.json`.
2. **Built Stateless `excel_parsing/pipeline.py` SDK:**
   - Provides `ExcelParser`, `inspect_workbook()`, and `extract_workbook()` accepting `bytes | bytearray | BinaryIO | Path | str`, structured `ProgressEvent` callbacks, and `fail_on_accounting` gating.
3. **Added CLI Entrypoint (`excel_parsing/cli.py` + `pyproject.toml` `[project.scripts]`):**
   - Provides `excel-parse inspect` and `excel-parse extract`.
4. **Verified Layer 1 (`tests/test_sdk.py`):**
   - Verified in-memory stream/bytes loading, SQL identifier quoting, `category_rate` join accuracy, accounting failure gates, and multi-sheet catalog reconciliation.

### Phase 2 — Layer 2: FastAPI Async Service & vLLM Queue Worker (Planned)
1. **Define API Schemas (`service/models.py`):**
   - Request/response Pydantic models for `/v1/inspect`, `/v1/jobs`, PostgreSQL target options, job progress snapshots, and webhook payloads.
2. **Implement Bounded Job Worker & Cache (`service/worker.py`):**
   - Single shared `VLLMClient` instance holding the global `max_concurrent_requests` semaphore across all active jobs.
   - Content-hash (`SHA-256`) idempotency store so duplicate uploads resolve immediately.
   - Per-region progress event emission for polling and SSE subscribers, plus optional webhook delivery on completion.
3. **Implement FastAPI App (`service/app.py`):**
   - Wire `POST /v1/inspect`, `POST /v1/jobs`, `GET /v1/jobs/{job_id}`, `GET /v1/jobs/{job_id}/events`, and `GET /healthz` / `/readyz`.
   - Enforce upload size limits, `.xlsx` MIME/magic-byte validation, and structured error responses.
4. **Containerization & End-to-End Verification:**
   - Add `Dockerfile` / `docker-compose.yml` profile configured against `VLLM_BASE_URL`.
   - Verify concurrent job submission, SSE progress streaming, idempotency cache hits, and executing the returned `pg_recordset` payload directly against PostgreSQL.
