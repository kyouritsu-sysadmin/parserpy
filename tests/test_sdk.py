"""Unit and integration tests for the Layer 1 `excel_parsing` Python SDK."""

from __future__ import annotations

import io
import json
from pathlib import Path
import tempfile
import unittest

import openpyxl

from excel_parsing import (
    ExcelParser,
    ImportCategory,
    ImportExtraction,
    ImportItem,
    InvalidWorkbookError,
    ParserConfig,
    RowAccountingError,
    SheetNotFoundError,
    build_pg_recordset_payload,
    denormalize_extraction,
    inspect_workbook,
)
from excel_parsing.cli import main as cli_main
from excel_parsing.reconcile import IgnoredRow, reconcile_saved_extraction, save_reconciliation

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLES_DIR = REPO_ROOT / "samples" / "catalogs"


def _build_in_memory_xlsx() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catalog"
    ws["A1"] = "No"
    ws["B1"] = "機器名"
    ws["C1"] = "hidden_col"
    ws["D1"] = "単価"
    ws.column_dimensions["C"].hidden = True

    ws.merge_cells("A2:D2")
    ws["A2"] = "特殊機器 40%"

    ws["A3"] = "A-01"
    ws["B3"] = "ACB 3P 630A"
    ws["C3"] = 999
    ws["D3"] = 1456030

    ws["A4"] = "A-02"
    ws["B4"] = "ELB 50A 4P無し"
    ws["C4"] = 999
    ws["D4"] = "-"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class _FakeVLLMClient:
    """Deterministic in-memory fake VLLMClient for fast unit testing of `ExcelParser.extract`."""

    def __init__(self, *, omit_row: int | None = None) -> None:
        self.omit_row = omit_row
        self.model = "fake-qwen"

    def resolve_model(self) -> str:
        return self.model

    def chat_json(self, system_prompt: str, user_prompt: str, *, schema_name: str = "response", **kwargs):
        if schema_name == "row_skeleton":
            return {
                "rows": [
                    {"row": 1, "role": "column_header"},
                    {"row": 2, "role": "category_header"},
                    {"row": 3, "role": "line_item"},
                    {"row": 4, "role": "line_item"},
                ]
            }
        records = []
        if self.omit_row != 3:
            records.append(
                {
                    "source_row": 3,
                    "name": "ACB 3P 630A",
                    "parent_row": 2,
                    "category": "特殊機器 40%",
                    "group": None,
                    "fields": {"No": "a-01", "単価": 1456030},
                }
            )
        if self.omit_row != 4:
            records.append(
                {
                    "source_row": 4,
                    "name": "ELB 50A 4P無し",
                    "parent_row": 2,
                    "category": "特殊機器 40%",
                    "group": None,
                    "fields": {"No": "a-02", "単価": "-"},
                }
            )
        return {"records": records}


class ExcelParsingSDKTests(unittest.TestCase):
    def test_input_sources_and_errors(self) -> None:
        raw = _build_in_memory_xlsx()

        # bytes, bytearray, BinaryIO all work without touching disk
        ins_bytes = inspect_workbook(raw)
        ins_ba = inspect_workbook(bytearray(raw))
        ins_stream = inspect_workbook(io.BytesIO(raw))

        self.assertEqual(ins_bytes.visible_sheet_names, ["Catalog"])
        self.assertEqual(ins_bytes.sheets[0].hidden_cols, 1)
        self.assertEqual(ins_ba.sheets[0].visible_cells, ins_stream.sheets[0].visible_cells)

        with self.assertRaises(InvalidWorkbookError):
            inspect_workbook(b"not a zip file")

        with self.assertRaises(SheetNotFoundError):
            inspect_workbook(raw, sheet_names=["DoesNotExist"])

    def test_pg_recordset_builder_variants(self) -> None:
        ext = ImportExtraction(
            categories=[
                ImportCategory(name="ACB(ZCT付) 40", rate=40.0),
                ImportCategory(name="【DS】", rate=None),
            ],
            items=[
                ImportItem(
                    source_row=6,
                    id="a-01",
                    name="ACB 3P 630A",
                    category="ACB(ZCT付) 40",
                    group="引出型",
                    unit_price=1456030,
                    note=None,
                ),
                ImportItem(
                    source_row=10,
                    id=None,
                    name="DS 400A",
                    category="【DS】",
                    group=None,
                    unit_price=43910,
                    note="標準品",
                ),
            ],
            ignored_rows=[IgnoredRow(source_row=5, reason="heading")],
        )

        rows = denormalize_extraction(ext)
        self.assertEqual(rows[0]["id"], "A-01")
        self.assertEqual(rows[0]["category_rate"], 40.0)
        self.assertIsNone(rows[1]["category_rate"])

        # Default asyncpg ($1::json) INSERT with reserved "group" identifier safely quoted
        p1 = build_pg_recordset_payload(ext, table="public.equipment_catalog")
        self.assertEqual(p1.row_count, 2)
        self.assertIn("INSERT INTO public.equipment_catalog", p1.query)
        self.assertIn("json_to_recordset($1::json) AS r(", p1.query)
        self.assertIn('"group" text', p1.query)
        self.assertEqual(len(json.loads(p1.params[0])), 2)

        # Custom column_map + extra_columns + ON CONFLICT + psycopg (%s::jsonb)
        p2 = build_pg_recordset_payload(
            ext,
            table="catalog_items",
            column_map={"id": "catalog_id", "group": "item_group"},
            extra_columns={"import_batch_id": "%s::uuid"},
            param_style="pyformat",
            json_type="jsonb",
            on_conflict="(catalog_id) DO UPDATE SET unit_price = EXCLUDED.unit_price",
            returning="catalog_id",
        )
        self.assertIn(
            "INSERT INTO catalog_items (import_batch_id, source_row, catalog_id, name, category, category_rate, item_group, unit_price, note)",
            p2.query,
        )
        self.assertIn("FROM jsonb_to_recordset(%s::jsonb)", p2.query)
        self.assertIn("ON CONFLICT (catalog_id) DO UPDATE", p2.query)
        self.assertIn("RETURNING catalog_id", p2.query)

        # Standalone SELECT mode + named SQLAlchemy parameter
        p3 = build_pg_recordset_payload(
            ext,
            mode="select",
            param_style="named",
            include_fields=("id", "name", "category_rate", "unit_price"),
        )
        self.assertTrue(p3.query.startswith("SELECT r.id, r.name, r.category_rate, r.unit_price FROM json_to_recordset(CAST(:records AS json))"))
        self.assertEqual(set(json.loads(p3.named_params["records"])[0].keys()), {"id", "name", "category_rate", "unit_price"})

    def test_end_to_end_fake_llm_and_accounting_gate(self) -> None:
        raw = _build_in_memory_xlsx()
        events = []

        parser_ok = ExcelParser(client=_FakeVLLMClient())  # type: ignore[arg-type]
        with tempfile.TemporaryDirectory() as tmp:
            res = parser_ok.extract(
                raw,
                fail_on_accounting=True,
                output_dir=tmp,
                pg_table="equipment_catalog",
                on_progress=events.append,
            )
            self.assertTrue(res.all_accounting_passed)
            self.assertEqual(res.total_valid_items, 2)
            self.assertEqual(len(events), 1)

            pg = res.to_pg_recordset(table="equipment_catalog")
            self.assertEqual(pg.row_count, 2)
            self.assertEqual(pg.records[0]["category_rate"], 40.0)
            self.assertIsNone(pg.records[1]["unit_price"])
            self.assertTrue((Path(tmp) / "extraction" / "Catalog" / "_pg_recordset.json").is_file())
            self.assertTrue((Path(tmp) / "extraction" / "_workbook_pg_recordset.json").is_file())

        # Verify fail_on_accounting raises RowAccountingError when a row remains missing
        parser_fail = ExcelParser(
            ParserConfig(retry_missing=False),
            client=_FakeVLLMClient(omit_row=4),  # type: ignore[arg-type]
        )
        with self.assertRaises(RowAccountingError):
            parser_fail.extract(raw, fail_on_accounting=True)

    def test_sample_inspection_and_saved_reconciliation_export(self) -> None:
        sample_file = SAMPLES_DIR / "●ツメ表ひな多島【面数換算】_1～20島_20260108.xlsx"
        if sample_file.is_file():
            rc = cli_main(["inspect", str(sample_file)])
            self.assertEqual(rc, 0)

        saved_root = REPO_ROOT / "outputs" / "extraction"
        for sheet_name in ("初回", "その他機器"):
            sheet_dir = saved_root / sheet_name
            if (sheet_dir / "_summary.json").is_file():
                sr = reconcile_saved_extraction(sheet_dir)
                files = save_reconciliation(sr, sheet_dir, pg_table="equipment_catalog")
                self.assertTrue(sr.all_accounting_passed)
                self.assertIsNotNone(sr.extraction)
                self.assertIn("pg_recordset", files)
                pg_json = json.loads(files["pg_recordset"].read_text(encoding="utf-8"))
                self.assertEqual(pg_json["row_count"], len(sr.extraction.items))
                self.assertGreater(pg_json["row_count"], 200)
                self.assertTrue(pg_json["all_accounting_passed"])


if __name__ == "__main__":
    unittest.main()
