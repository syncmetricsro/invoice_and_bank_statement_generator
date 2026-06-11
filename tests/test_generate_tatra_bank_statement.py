from __future__ import annotations

import csv
import datetime as dt
import importlib.util
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
INVOICE_MODULE_PATH = ROOT / "scripts" / "bulk_zalohova_faktura_generator.py"
STATEMENT_MODULE_PATH = ROOT / "scripts" / "generate_tatra_bank_statement.py"

INVOICE_SPEC = importlib.util.spec_from_file_location("bulk_generator", INVOICE_MODULE_PATH)
INVOICE_MODULE = importlib.util.module_from_spec(INVOICE_SPEC)
assert INVOICE_SPEC and INVOICE_SPEC.loader
sys.modules[INVOICE_SPEC.name] = INVOICE_MODULE
INVOICE_SPEC.loader.exec_module(INVOICE_MODULE)

STATEMENT_SPEC = importlib.util.spec_from_file_location("bank_statement_generator", STATEMENT_MODULE_PATH)
STATEMENT_MODULE = importlib.util.module_from_spec(STATEMENT_SPEC)
assert STATEMENT_SPEC and STATEMENT_SPEC.loader
sys.modules[STATEMENT_SPEC.name] = STATEMENT_MODULE
STATEMENT_SPEC.loader.exec_module(STATEMENT_MODULE)


def write_invoice_manifest(path: Path, count: int = 1000) -> None:
    records = INVOICE_MODULE.build_batch_plan(
        count=count,
        start_index=1,
        issue_date=dt.date(2026, 5, 26),
        seed=42,
        customers_csv=None,
    )
    rows = [INVOICE_MODULE.invoice_manifest_row(record, include_pdf=True) for record in records]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


class GenerateTatraBankStatementTests(unittest.TestCase):
    def test_missing_required_columns_fail_clearly(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            bad_csv = Path(tmp_dir) / "bad.csv"
            bad_csv.write_text("customer_id,variable_symbol\nA,1\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "missing required columns"):
                STATEMENT_MODULE.read_invoice_rows(bad_csv)

    def test_parse_supplier_iban_into_statement_fields(self) -> None:
        bank_id, account_id = STATEMENT_MODULE.parse_sk_iban("SK1211000000002987654321")
        self.assertEqual(bank_id, "1100")
        self.assertEqual(account_id, "2987654321")

    def test_statement_generation_is_deterministic(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            invoices_csv = tmp / "invoices.csv"
            write_invoice_manifest(invoices_csv, count=25)
            rows = STATEMENT_MODULE.read_invoice_rows(invoices_csv)

            first = STATEMENT_MODULE.generate_statement(rows, seed=42)
            second = STATEMENT_MODULE.generate_statement(rows, seed=42)

            self.assertEqual(first[0], second[0])
            self.assertEqual(first[1], second[1])
            self.assertEqual(first[2], second[2])
            self.assertEqual(first[3], second[3])

    def test_scenario_coverage_for_1000_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            invoices_csv = Path(tmp_dir) / "invoices.csv"
            write_invoice_manifest(invoices_csv, count=1000)
            rows = STATEMENT_MODULE.read_invoice_rows(invoices_csv)
            transactions, tx_expectations, reconciliation_expectations, summary = STATEMENT_MODULE.generate_statement(rows, seed=42)

            reasons = {row.expected_reason: 0 for row in reconciliation_expectations}
            for row in reconciliation_expectations:
                reasons[row.expected_reason] = reasons.get(row.expected_reason, 0) + 1

            # 1000 monthly invoices plus 200 annual extras (every 5th customer),
            # all extras exact_single: 600 exacts, every 10th becomes a duplicate.
            self.assertEqual(len(rows), 1200)
            self.assertEqual(reasons["exact_match"], 540)
            self.assertEqual(reasons["duplicate_possible"], 60)
            self.assertEqual(reasons["split_payment_possible"], 200)
            self.assertEqual(reasons["underpayment"], 200)
            self.assertEqual(reasons["overpayment"], 200)
            self.assertEqual(summary["noise_credit_count"], 24)
            self.assertEqual(summary["noise_debit_count"], 24)
            self.assertEqual(len(tx_expectations), len(transactions))

    def test_transaction_integrity(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            invoices_csv = Path(tmp_dir) / "invoices.csv"
            write_invoice_manifest(invoices_csv, count=1000)
            rows = STATEMENT_MODULE.read_invoice_rows(invoices_csv)
            transactions, tx_expectations, reconciliation_expectations, _summary = STATEMENT_MODULE.generate_statement(rows, seed=42)

            ids = [transaction.transaction_id for transaction in transactions]
            self.assertEqual(len(ids), len(set(ids)))

            real_vs = {row.variable_symbol for row in rows}
            for expectation, transaction in zip(sorted(tx_expectations, key=lambda item: item.transaction_id), sorted(transactions, key=lambda item: item.transaction_id)):
                if expectation.expected_category == "unknown_credit":
                    self.assertNotIn(transaction.variable_symbol, real_vs)

            split_example = next(row for row in reconciliation_expectations if row.expected_reason == "split_payment_possible")
            split_invoice = next(invoice for invoice in rows if invoice.variable_symbol == split_example.variable_symbol)
            self.assertEqual(Decimal(split_example.expected_received_total), sum(split_invoice.simulated_split_amounts, Decimal("0.00")))

            duplicate_example = next(row for row in reconciliation_expectations if row.expected_reason == "duplicate_possible")
            duplicate_ids = duplicate_example.expected_matched_transaction_ids.split(",")
            duplicate_transactions = [transaction for transaction in transactions if transaction.transaction_id in duplicate_ids]
            self.assertEqual(len(duplicate_transactions), 2)
            self.assertEqual(duplicate_transactions[0].amount, duplicate_transactions[1].amount)
            self.assertEqual(Decimal(duplicate_transactions[0].amount) * 2, Decimal(duplicate_example.expected_received_total))

            annual_extra_example = next(row for row in reconciliation_expectations if row.invoice_no.endswith("-AE"))
            matched_ids = annual_extra_example.expected_matched_transaction_ids.split(",")
            matched_transactions = [transaction for transaction in transactions if transaction.transaction_id in matched_ids]
            self.assertEqual(len(matched_transactions), int(annual_extra_example.expected_matched_transaction_count))
            self.assertEqual(
                sum(Decimal(transaction.amount) for transaction in matched_transactions),
                Decimal(annual_extra_example.expected_received_total),
            )

    def test_credit_booking_dates_stay_inside_matching_window(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            invoices_csv = Path(tmp_dir) / "invoices.csv"
            write_invoice_manifest(invoices_csv, count=100)
            rows = STATEMENT_MODULE.read_invoice_rows(invoices_csv)
            transactions, _tx_expectations, _recon, _summary = STATEMENT_MODULE.generate_statement(rows, seed=42)
            by_vs = {row.variable_symbol: row for row in rows}

            for transaction in transactions:
                if transaction.direction != "credit" or transaction.variable_symbol not in by_vs:
                    continue
                invoice = by_vs[transaction.variable_symbol]
                booking = dt.datetime.strptime(transaction.booking_date, "%Y-%m-%d").date()
                start, end = STATEMENT_MODULE.payment_window(invoice)
                self.assertGreaterEqual(booking, start)
                self.assertLessEqual(booking, end)

    def test_ofx_count_matches_transactions_and_is_well_formed(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            invoices_csv = tmp / "invoices.csv"
            write_invoice_manifest(invoices_csv, count=50)
            rows = STATEMENT_MODULE.read_invoice_rows(invoices_csv)
            transactions, _tx_expectations, _recon, summary = STATEMENT_MODULE.generate_statement(rows, seed=42)

            ofx_path = tmp / "statement.ofx"
            STATEMENT_MODULE.write_ofx(ofx_path, transactions, summary)

            tree = ET.parse(ofx_path)
            root = tree.getroot()
            entries = root.findall(".//STMTTRN")
            self.assertEqual(len(entries), len(transactions))
            first_entry = entries[0]
            self.assertIsNotNone(first_entry.find("FITID"))
            self.assertIsNotNone(first_entry.find("REFERENCE_E2E"))
            self.assertIsNotNone(first_entry.find("CURRENCY"))

    def test_cli_writes_standard_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            invoices_csv = tmp / "invoices.csv"
            outdir = tmp / "bank_out"
            write_invoice_manifest(invoices_csv, count=25)

            # Exercise the CLI entrypoint through a subprocess-like direct call pattern.
            import subprocess

            result = subprocess.run(
                [
                    sys.executable,
                    str(STATEMENT_MODULE_PATH),
                    "--invoices",
                    str(invoices_csv),
                    "--outdir",
                    str(outdir),
                    "--seed",
                    "42",
                ],
                check=True,
                capture_output=True,
                text=True,
            )
            self.assertIn("Generated OFX statement", result.stdout)
            for filename in (
                "transactions.csv",
                "statement.ofx",
                "transaction_expectations.csv",
                "transaction_expectations.json",
                "reconciliation_expectations.csv",
                "reconciliation_expectations.json",
                "summary.json",
            ):
                self.assertTrue((outdir / filename).exists(), filename)


if __name__ == "__main__":
    unittest.main()
