from __future__ import annotations

import csv
import datetime as dt
import importlib.util
import sys
import tempfile
import unittest
from decimal import Decimal
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

        self.assertEqual(len(records), 1200)
        self.assertEqual(sum(record.amount_bucket == "random_whole_eur" for record in records), 150)
        self.assertEqual(sum(record.payment_scenario == "exact_split_total" for record in records), 200)

        monthly_records = [record for record in records if record.charge_type == "monthly"]
        annual_extras = [record for record in records if record.charge_type == "annual_extra"]
        self.assertEqual(len(monthly_records), 1000)
        self.assertEqual(len(annual_extras), 200)

        monthly_customer_ids = {record.invoice.customer.customer_id for record in monthly_records}
        for record in annual_extras:
            invoice = record.invoice
            self.assertEqual(invoice.total, MODULE.ANNUAL_EXTRA_FEE)
            self.assertEqual(record.payment_scenario, "exact_single")
            self.assertEqual(record.amount_bucket, "annual_extra")
            self.assertTrue(invoice.invoice_no.endswith("-AE"))
            self.assertEqual(len(invoice.variable_symbol), 9)
            self.assertIn(invoice.customer.customer_id, monthly_customer_ids)

        variable_symbols = [record.invoice.variable_symbol for record in records]
        self.assertEqual(len(variable_symbols), len(set(variable_symbols)))

        for record in records:
            invoice = record.invoice
            self.assertEqual(invoice.total, invoice.unit_price)
            if record.amount_bucket == "random_whole_eur":
                self.assertNotIn(int(invoice.total), {80, 180, 210})
            if record.payment_scenario == "exact_split_total":
                self.assertEqual(sum(record.simulated_split_amounts), record.simulated_paid_total)
                self.assertEqual(len(record.simulated_split_amounts), 2)
            if record.payment_scenario == "underpay":
                self.assertLess(record.simulated_paid_total, invoice.total)
            if record.payment_scenario == "overpay":
                self.assertGreater(record.simulated_paid_total, invoice.total)

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

    def test_expected_charge_manifest_uses_owed_amount_and_due_date(self) -> None:
        records = MODULE.build_batch_plan(
            count=5,
            start_index=1,
            issue_date=dt.date(2026, 5, 25),
            seed=42,
            customers_csv=None,
        )

        underpay_record = next(record for record in records if record.payment_scenario == "underpay")
        row = MODULE.expected_charge_manifest_row(underpay_record)

        self.assertEqual(row["customer_id"], underpay_record.invoice.customer.customer_id)
        self.assertEqual(row["billing_month"], "2026-05")
        self.assertEqual(row["charge_type"], "monthly")
        self.assertEqual(row["variable_symbol"], underpay_record.invoice.variable_symbol)
        self.assertEqual(row["charge_amount"], "210.00")
        self.assertEqual(row["due_date"], underpay_record.invoice.due_date.isoformat())
        self.assertNotEqual(row["charge_amount"], MODULE.decimal_string(underpay_record.simulated_paid_total))

    def test_customer_billing_profile_is_deterministic(self) -> None:
        records = MODULE.build_batch_plan(
            count=10,
            start_index=1,
            issue_date=dt.date(2026, 5, 25),
            seed=42,
            customers_csv=None,
        )

        monthly_records = [record for record in records if record.charge_type == "monthly"]
        annual_extras = {record.invoice.customer.customer_id: record for record in records if record.charge_type == "annual_extra"}
        self.assertEqual(len(monthly_records), 10)
        self.assertEqual(set(annual_extras), {"CUST-00005", "CUST-00010"})

        for offset, record in enumerate(monthly_records):
            index = 1 + offset
            customer = record.invoice.customer
            self.assertEqual(customer.monthly_fee, MODULE.decimal_string(record.invoice.total))
            self.assertEqual(customer.vat_status, "vat_payer" if customer.ic_dph else "non_vat_payer")
            self.assertEqual(customer.payment_method, MODULE.customer_payment_method(index))
            self.assertEqual(customer.status, MODULE.customer_status(index))
            if MODULE.has_annual_extra(index):
                self.assertEqual(customer.annual_extra_fee, MODULE.decimal_string(MODULE.ANNUAL_EXTRA_FEE))
                self.assertEqual(customer.annual_extra_interval_months, str(MODULE.ANNUAL_EXTRA_INTERVAL_MONTHS))
                extra = annual_extras[customer.customer_id]
                self.assertEqual(MODULE.decimal_string(extra.invoice.total), customer.annual_extra_fee)
            else:
                self.assertEqual(customer.annual_extra_fee, "")
                self.assertEqual(customer.annual_extra_interval_months, "")

        manifest_row = MODULE.customer_manifest_row(monthly_records[0].invoice.customer)
        for key in ("monthly_fee", "annual_extra_fee", "annual_extra_interval_months", "payment_method", "vat_status", "status"):
            self.assertIn(key, manifest_row)

    def test_months_one_keeps_plain_invoice_number(self) -> None:
        records = MODULE.build_batch_plan(
            count=2,
            start_index=1,
            issue_date=dt.date(2026, 6, 14),
            seed=42,
            customers_csv=None,
        )
        # Single-month invoice numbers carry no -MM suffix (unchanged behavior).
        for record in records:
            self.assertRegex(record.invoice.invoice_no, r"^\d{4}-\d{4}$")

    def test_months_multi_emits_recurring_series(self) -> None:
        records = MODULE.build_batch_plan(
            count=1,
            start_index=1,
            issue_date=dt.date(2026, 1, 14),
            seed=42,
            customers_csv=None,
            months=6,
        )
        self.assertEqual(len(records), 6)

        # Same VS and same amount every month; distinct, month-suffixed numbers.
        self.assertEqual({r.invoice.variable_symbol for r in records}, {"20260001"})
        self.assertEqual({MODULE.decimal_string(r.invoice.total) for r in records}, {MODULE.decimal_string(records[0].invoice.total)})
        self.assertEqual(
            [r.invoice.invoice_no for r in records],
            [f"2026-0001-{m:02d}" for m in range(1, 7)],
        )
        self.assertEqual(len({r.filename_base for r in records}), 6)
        self.assertEqual(
            [MODULE.billing_month(r.invoice.issue_date) for r in records],
            ["2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06"],
        )

        # First five paid, final month unpaid.
        self.assertTrue(all(r.payment_scenario == "exact_single" for r in records[:5]))
        self.assertEqual(records[5].payment_scenario, "unpaid")
        self.assertEqual(records[5].simulated_paid_total, Decimal("0.00"))

    def test_write_manifests_creates_expected_charges_files(self) -> None:
        records = MODULE.build_batch_plan(
            count=3,
            start_index=1,
            issue_date=dt.date(2026, 5, 25),
            seed=42,
            customers_csv=None,
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            MODULE.write_manifests(
                tmp_path,
                records,
                include_pdf=True,
                template=ROOT / "docs" / "zalohova_faktura_template_ready.docx",
                seed=42,
            )

            expected_csv = tmp_path / "manifests" / "expected_charges.csv"
            expected_json = tmp_path / "manifests" / "expected_charges.json"
            self.assertTrue(expected_csv.exists())
            self.assertTrue(expected_json.exists())

            with (tmp_path / "manifests" / "customers.csv").open("r", encoding="utf-8", newline="") as handle:
                header = next(csv.reader(handle))
            for column in ("monthly_fee", "annual_extra_fee", "annual_extra_interval_months", "payment_method", "vat_status", "status"):
                self.assertIn(column, header)


if __name__ == "__main__":
    unittest.main()
