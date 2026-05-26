from __future__ import annotations

import csv
import datetime as dt
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "bulk_zalohova_faktura_generator.py"
SPEC = importlib.util.spec_from_file_location("bulk_generator", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class BulkGeneratorTests(unittest.TestCase):
    def test_amount_bucket_distribution_for_1000(self) -> None:
        sequence = MODULE.build_weighted_sequence(1000, MODULE.AMOUNT_BUCKET_SPECS)
        self.assertEqual(sequence.count("fixed_80"), 300)
        self.assertEqual(sequence.count("fixed_180"), 300)
        self.assertEqual(sequence.count("fixed_210"), 250)
        self.assertEqual(sequence.count("random_whole_eur"), 150)

    def test_payment_scenario_distribution_for_1000(self) -> None:
        sequence = MODULE.build_weighted_sequence(1000, MODULE.PAYMENT_SCENARIO_SPECS)
        self.assertEqual(sequence.count("exact_single"), 400)
        self.assertEqual(sequence.count("exact_split_total"), 200)
        self.assertEqual(sequence.count("underpay"), 200)
        self.assertEqual(sequence.count("overpay"), 200)

    def test_generated_customers_are_unique(self) -> None:
        customers = [MODULE.generated_customer(index, seed=42) for index in range(1, 1001)]
        self.assertEqual(len({customer.customer_id for customer in customers}), 1000)
        self.assertEqual(len({customer.name for customer in customers}), 1000)
        self.assertEqual(len({(customer.street, customer.city_country) for customer in customers}), 1000)
        self.assertEqual(len({customer.iban_compact for customer in customers}), 1000)

    def test_partial_csv_is_backfilled_and_enriched(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_path = Path(tmp_dir) / "customers.csv"
            with csv_path.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.DictWriter(
                    handle,
                    fieldnames=["customer_name", "street", "city_country", "customer_id", "email"],
                )
                writer.writeheader()
                writer.writerow({
                    "customer_name": "CSV Customer 1 s.r.o.",
                    "street": "CSV Street 1",
                    "city_country": "811 01 Bratislava, Slovenská republika",
                    "customer_id": "CSV-001",
                    "email": "",
                })
                writer.writerow({
                    "customer_name": "CSV Customer 2 s.r.o.",
                    "street": "",
                    "city_country": "",
                    "customer_id": "",
                    "email": "csv2@example.test",
                })

            customers = MODULE.resolve_customers(
                count=5,
                start_index=1,
                seed=42,
                customers_csv=csv_path,
            )

            self.assertEqual(len(customers), 5)
            self.assertEqual(customers[0].name, "CSV Customer 1 s.r.o.")
            self.assertEqual(customers[0].customer_id, "CSV-001")
            self.assertTrue(customers[0].email.endswith("@example.test"))
            self.assertEqual(customers[1].name, "CSV Customer 2 s.r.o.")
            self.assertTrue(customers[1].street)
            self.assertTrue(customers[1].city_country)
            self.assertTrue(customers[1].customer_id)
            self.assertEqual(customers[1].email, "csv2@example.test")
            self.assertTrue(customers[2].name)

    def test_planned_batch_contains_consistent_manifest_metadata(self) -> None:
        records = MODULE.build_batch_plan(
            count=1000,
            start_index=1,
            issue_date=dt.date(2026, 5, 25),
            seed=42,
            customers_csv=None,
        )

        self.assertEqual(len(records), 1000)
        self.assertEqual(sum(record.amount_bucket == "random_whole_eur" for record in records), 150)
        self.assertEqual(sum(record.payment_scenario == "exact_split_total" for record in records), 200)

        for record in records:
            invoice = record.invoice
            self.assertEqual(invoice.total, invoice.unit_price)
            if record.amount_bucket == "random_whole_eur":
                self.assertNotIn(int(invoice.total), {80, 180, 210})
            if record.payment_scenario == "exact_split_total":
                self.assertEqual(sum(record.suggested_split_amounts), record.suggested_received_amount)
                self.assertEqual(len(record.suggested_split_amounts), 2)
            if record.payment_scenario == "underpay":
                self.assertLess(record.suggested_received_amount, invoice.total)
            if record.payment_scenario == "overpay":
                self.assertGreater(record.suggested_received_amount, invoice.total)

    def test_direct_pdf_renderer_writes_pdf(self) -> None:
        record = MODULE.build_batch_plan(
            count=1,
            start_index=1,
            issue_date=dt.date(2026, 5, 25),
            seed=42,
            customers_csv=None,
        )[0]

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            qr_path = tmp_path / "qr.png"
            pdf_path = tmp_path / "invoice.pdf"
            MODULE.create_qr_png(record.invoice, qr_path)
            MODULE.render_invoice_pdf(pdf_path, record.invoice, qr_path)
            self.assertTrue(pdf_path.exists())
            self.assertGreater(pdf_path.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
