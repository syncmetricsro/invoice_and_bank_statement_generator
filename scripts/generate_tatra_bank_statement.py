#!/usr/bin/env python3
"""
Generate deterministic Tatra-style bank statement test data from invoice manifests.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
from dataclasses import asdict, dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


REQUIRED_COLUMNS = [
    "batch_id",
    "customer_id",
    "customer_name",
    "variable_symbol",
    "billing_month",
    "issue_date",
    "due_date",
    "invoice_total_amount",
    "payment_scenario",
    "simulated_paid_total",
    "simulated_split_amounts",
    "supplier_name",
    "supplier_iban",
    "customer_iban",
    "customer_bank_code",
    "customer_account_number",
    "reference_text_template",
]

UNKNOWN_CREDIT_PATTERNS = (
    ("missing_vs", "Monthly subscription without symbol", ""),
    ("unknown_vs", "/VS91999999/KS0308/TXT Unknown payer monthly service", "91999999"),
    ("unknown_vs", "VS: 92999999 Follow-up service for special customer", "92999999"),
)

VENDOR_DEBIT_PATTERNS = (
    ("Hotel Tatra Bratislava", "HOTEL TATRA BOOKING business trip"),
    ("PrintCity", "PRINT SHOP BUSINESS CARDS"),
    ("Alza.sk", "ALZA.SK OFFICE SUPPLIES"),
    ("Bolt Business", "BOLT BUSINESS RIDES"),
)


@dataclass(frozen=True)
class InvoiceRow:
    batch_id: str
    customer_id: str
    customer_name: str
    variable_symbol: str
    billing_month: str
    issue_date: dt.date
    due_date: dt.date
    invoice_total_amount: Decimal
    payment_scenario: str
    simulated_paid_total: Decimal
    simulated_split_amounts: tuple[Decimal, ...]
    supplier_name: str
    supplier_iban: str
    customer_iban: str
    customer_bank_code: str
    customer_account_number: str
    reference_text_template: str
    invoice_no: str


@dataclass(frozen=True)
class BankTransaction:
    import_batch_id: str
    source: str
    transaction_id: str
    booking_date: str
    value_date: str
    amount: str
    currency: str
    direction: str
    payer_name: str
    counterparty_iban: str
    reference_text: str
    variable_symbol: str
    specific_symbol: str
    constant_symbol: str
    raw_description: str
    raw_source_reference: str
    imported_at: str


@dataclass(frozen=True)
class TransactionExpectation:
    transaction_id: str
    expected_category: str
    expected_customer_id: str
    expected_variable_symbol: str
    expected_direction: str
    notes: str


@dataclass(frozen=True)
class ReconciliationExpectation:
    customer_id: str
    invoice_no: str
    variable_symbol: str
    expected_status: str
    expected_reason: str
    expected_invoice_action: str
    expected_received_total: str
    expected_matched_transaction_count: str
    expected_matched_transaction_ids: str


def money(value: Decimal | float | int | str) -> Decimal:
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def decimal_string(value: Decimal | float | int | str) -> str:
    return f"{money(value):.2f}"


def compact_iban(value: str) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()


def parse_sk_iban(value: str) -> tuple[str, str]:
    compact = compact_iban(value)
    if not re.fullmatch(r"SK\d{22}", compact):
        raise ValueError(f"Unsupported IBAN format: {value}")
    bban = compact[4:]
    return bban[:4], bban[-10:]


def parse_date(value: str) -> dt.date:
    return dt.datetime.strptime(value, "%Y-%m-%d").date()


def parse_split_amounts(value: str) -> tuple[Decimal, ...]:
    if not value.strip():
        return ()
    raw_values = json.loads(value)
    return tuple(money(raw) for raw in raw_values)


def require_columns(fieldnames: list[str] | None) -> list[str]:
    if not fieldnames:
        raise ValueError("Input CSV has no header row")
    missing = [column for column in REQUIRED_COLUMNS if column not in fieldnames]
    if missing:
        raise ValueError(f"invoices.csv is missing required columns: {', '.join(missing)}")
    return fieldnames


def read_invoice_rows(path: Path) -> list[InvoiceRow]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        require_columns(reader.fieldnames)
        rows: list[InvoiceRow] = []
        for raw in reader:
            rows.append(InvoiceRow(
                batch_id=str(raw["batch_id"]).strip(),
                customer_id=str(raw["customer_id"]).strip(),
                customer_name=str(raw["customer_name"]).strip(),
                variable_symbol=str(raw["variable_symbol"]).strip(),
                billing_month=str(raw["billing_month"]).strip(),
                issue_date=parse_date(str(raw["issue_date"]).strip()),
                due_date=parse_date(str(raw["due_date"]).strip()),
                invoice_total_amount=money(raw["invoice_total_amount"]),
                payment_scenario=str(raw["payment_scenario"]).strip(),
                simulated_paid_total=money(raw["simulated_paid_total"]),
                simulated_split_amounts=parse_split_amounts(str(raw["simulated_split_amounts"])),
                supplier_name=str(raw["supplier_name"]).strip(),
                supplier_iban=compact_iban(str(raw["supplier_iban"]).strip()),
                customer_iban=compact_iban(str(raw["customer_iban"]).strip()),
                customer_bank_code=str(raw["customer_bank_code"]).strip(),
                customer_account_number=str(raw["customer_account_number"]).strip(),
                reference_text_template=str(raw["reference_text_template"]).strip(),
                invoice_no=str(raw["invoice_no"]).strip(),
            ))
        if not rows:
            raise ValueError(f"No invoice rows found in {path}")
        return rows


def month_end(date_value: dt.date) -> dt.date:
    next_month = (date_value.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
    return next_month - dt.timedelta(days=1)


def payment_window(invoice: InvoiceRow) -> tuple[dt.date, dt.date]:
    start = invoice.due_date - dt.timedelta(days=4)
    month_limit = month_end(invoice.issue_date) + dt.timedelta(days=14)
    end = min(invoice.due_date + dt.timedelta(days=5), month_limit)
    if end < start:
        end = start
    return start, end


def booking_date_for_invoice(invoice: InvoiceRow, index: int, seed: int) -> dt.date:
    start, end = payment_window(invoice)
    span = (end - start).days + 1
    offset = (index + seed) % max(span, 1)
    return start + dt.timedelta(days=offset)


def clamp_date(value: dt.date, end: dt.date) -> dt.date:
    return value if value <= end else end


def reference_variant(reference: str, variant: int) -> str:
    if variant % 3 == 0:
        return reference
    variable_symbol = extract_variable_symbol(reference)
    if variant % 3 == 1 and variable_symbol:
        month_match = re.search(r"(\d{4}-\d{2})", reference)
        suffix = month_match.group(1) if month_match else "billing cycle"
        return f"VS: {variable_symbol} Accounting services {suffix}"
    if variable_symbol:
        return reference.replace("/VS", "/vs", 1)
    return reference


def extract_variable_symbol(reference: str) -> str:
    match = re.search(r"(?:/VS|VS: ?|VS)(\d+)", reference, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def build_import_batch_id(batch_id: str, seed: int) -> str:
    return f"generated_tatra_ofx_{batch_id}_seed{seed}"


def build_imported_at(invoices: list[InvoiceRow]) -> str:
    latest_due = max(invoice.due_date for invoice in invoices)
    imported = dt.datetime.combine(latest_due + dt.timedelta(days=1), dt.time(12, 0))
    return imported.strftime("%Y-%m-%dT%H:%M:%S.000Z")


def build_transaction_id(prefix: str, number: int) -> str:
    return f"{prefix}-{number:05d}"


def noise_count(invoice_count: int) -> int:
    return max(2, invoice_count // 50)


def is_duplicate_exact(invoice: InvoiceRow, exact_counter: int) -> bool:
    return invoice.payment_scenario == "exact_single" and exact_counter % 10 == 0


def transaction_row(
    *,
    import_batch_id: str,
    transaction_id: str,
    booking_date: dt.date,
    amount: Decimal,
    direction: str,
    payer_name: str,
    counterparty_iban: str,
    reference_text: str,
    variable_symbol: str,
    raw_description: str,
    raw_source_reference: dict[str, Any],
    imported_at: str,
    constant_symbol: str = "",
) -> BankTransaction:
    return BankTransaction(
        import_batch_id=import_batch_id,
        source="generated_tatra_ofx",
        transaction_id=transaction_id,
        booking_date=booking_date.isoformat(),
        value_date=booking_date.isoformat(),
        amount=decimal_string(amount),
        currency="EUR",
        direction=direction,
        payer_name=payer_name,
        counterparty_iban=counterparty_iban,
        reference_text=reference_text,
        variable_symbol=variable_symbol,
        specific_symbol="",
        constant_symbol=constant_symbol,
        raw_description=raw_description,
        raw_source_reference=json.dumps(raw_source_reference, ensure_ascii=False, sort_keys=True),
        imported_at=imported_at,
    )


def csv_rows(items: list[Any]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for item in items:
        row: dict[str, str] = {}
        for key, value in asdict(item).items():
            row[key] = str(value)
        rows.append(row)
    return rows


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    if not rows:
        raise ValueError(f"No rows to write to {path}")
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def write_json(path: Path, payload: Any) -> None:
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=False)


def generate_statement(
    invoices: list[InvoiceRow],
    *,
    seed: int,
) -> tuple[list[BankTransaction], list[TransactionExpectation], list[ReconciliationExpectation], dict[str, Any]]:
    supplier_bank_id, supplier_account_id = parse_sk_iban(invoices[0].supplier_iban)
    import_batch_id = build_import_batch_id(invoices[0].batch_id, seed)
    imported_at = build_imported_at(invoices)
    transactions: list[BankTransaction] = []
    tx_expectations: list[TransactionExpectation] = []
    reconciliation_expectations: list[ReconciliationExpectation] = []
    exact_counter = 0
    transaction_number = 1
    real_variable_symbols = {invoice.variable_symbol for invoice in invoices}

    for invoice_index, invoice in enumerate(invoices, start=1):
        booking = booking_date_for_invoice(invoice, invoice_index, seed)
        _, window_end = payment_window(invoice)
        matched_ids: list[str] = []
        expected_status = "Paid"
        expected_reason = "exact_match"
        expected_action = "create_invoice"
        expected_received = invoice.invoice_total_amount
        category = "customer_payment"

        if invoice.payment_scenario == "exact_single":
            exact_counter += 1
            if is_duplicate_exact(invoice, exact_counter):
                expected_status = "Mismatch"
                expected_reason = "duplicate_possible"
                expected_action = "manual_review"
                expected_received = money(invoice.invoice_total_amount * 2)
                category = "duplicate_payment"
                for duplicate_index in range(2):
                    tx_id = build_transaction_id("TX", transaction_number)
                    transaction_number += 1
                    duplicate_date = clamp_date(booking + dt.timedelta(days=duplicate_index), window_end)
                    reference_text = reference_variant(invoice.reference_text_template, duplicate_index + 1)
                    transactions.append(transaction_row(
                        import_batch_id=import_batch_id,
                        transaction_id=tx_id,
                        booking_date=duplicate_date,
                        amount=invoice.invoice_total_amount,
                        direction="credit",
                        payer_name=invoice.customer_name,
                        counterparty_iban=invoice.customer_iban,
                        reference_text=reference_text,
                        variable_symbol=invoice.variable_symbol,
                        raw_description=reference_text,
                        raw_source_reference={
                            "customer_id": invoice.customer_id,
                            "invoice_no": invoice.invoice_no,
                            "scenario": "duplicate_payment",
                            "payment_index": duplicate_index + 1,
                        },
                        imported_at=imported_at,
                        constant_symbol="0308",
                    ))
                    tx_expectations.append(TransactionExpectation(
                        transaction_id=tx_id,
                        expected_category="duplicate_payment",
                        expected_customer_id=invoice.customer_id,
                        expected_variable_symbol=invoice.variable_symbol,
                        expected_direction="credit",
                        notes=f"Duplicate exact payment {duplicate_index + 1} for {invoice.invoice_no}",
                    ))
                    matched_ids.append(tx_id)
            else:
                tx_id = build_transaction_id("TX", transaction_number)
                transaction_number += 1
                transactions.append(transaction_row(
                    import_batch_id=import_batch_id,
                    transaction_id=tx_id,
                    booking_date=booking,
                    amount=invoice.invoice_total_amount,
                    direction="credit",
                    payer_name=invoice.customer_name,
                    counterparty_iban=invoice.customer_iban,
                    reference_text=invoice.reference_text_template,
                    variable_symbol=invoice.variable_symbol,
                    raw_description=invoice.reference_text_template,
                    raw_source_reference={
                        "customer_id": invoice.customer_id,
                        "invoice_no": invoice.invoice_no,
                        "scenario": invoice.payment_scenario,
                    },
                    imported_at=imported_at,
                    constant_symbol="0308",
                ))
                tx_expectations.append(TransactionExpectation(
                    transaction_id=tx_id,
                    expected_category=category,
                    expected_customer_id=invoice.customer_id,
                    expected_variable_symbol=invoice.variable_symbol,
                    expected_direction="credit",
                    notes=f"Exact payment for {invoice.invoice_no}",
                ))
                matched_ids.append(tx_id)

        elif invoice.payment_scenario == "exact_split_total":
            expected_status = "Mismatch"
            expected_reason = "split_payment_possible"
            expected_action = "manual_review"
            expected_received = sum(invoice.simulated_split_amounts, Decimal("0.00"))
            for split_index, amount in enumerate(invoice.simulated_split_amounts, start=1):
                tx_id = build_transaction_id("TX", transaction_number)
                transaction_number += 1
                split_date = clamp_date(booking + dt.timedelta(days=split_index - 1), window_end)
                reference_text = reference_variant(invoice.reference_text_template, split_index)
                transactions.append(transaction_row(
                    import_batch_id=import_batch_id,
                    transaction_id=tx_id,
                    booking_date=split_date,
                    amount=amount,
                    direction="credit",
                    payer_name=invoice.customer_name,
                    counterparty_iban=invoice.customer_iban,
                    reference_text=reference_text,
                    variable_symbol=invoice.variable_symbol,
                    raw_description=reference_text,
                    raw_source_reference={
                        "customer_id": invoice.customer_id,
                        "invoice_no": invoice.invoice_no,
                        "scenario": invoice.payment_scenario,
                        "payment_index": split_index,
                    },
                    imported_at=imported_at,
                    constant_symbol="0308",
                ))
                tx_expectations.append(TransactionExpectation(
                    transaction_id=tx_id,
                    expected_category="split_payment_component",
                    expected_customer_id=invoice.customer_id,
                    expected_variable_symbol=invoice.variable_symbol,
                    expected_direction="credit",
                    notes=f"Split payment component {split_index} for {invoice.invoice_no}",
                ))
                matched_ids.append(tx_id)

        elif invoice.payment_scenario in {"underpay", "overpay"}:
            tx_id = build_transaction_id("TX", transaction_number)
            transaction_number += 1
            expected_status = "Mismatch"
            expected_reason = "underpayment" if invoice.payment_scenario == "underpay" else "overpayment"
            expected_action = "manual_review"
            expected_received = invoice.simulated_paid_total
            reference_text = reference_variant(invoice.reference_text_template, invoice_index)
            transactions.append(transaction_row(
                import_batch_id=import_batch_id,
                transaction_id=tx_id,
                booking_date=booking,
                amount=invoice.simulated_paid_total,
                direction="credit",
                payer_name=invoice.customer_name,
                counterparty_iban=invoice.customer_iban,
                reference_text=reference_text,
                variable_symbol=invoice.variable_symbol,
                raw_description=reference_text,
                raw_source_reference={
                    "customer_id": invoice.customer_id,
                    "invoice_no": invoice.invoice_no,
                    "scenario": invoice.payment_scenario,
                },
                imported_at=imported_at,
                constant_symbol="0308",
            ))
            tx_expectations.append(TransactionExpectation(
                transaction_id=tx_id,
                expected_category="customer_payment",
                expected_customer_id=invoice.customer_id,
                expected_variable_symbol=invoice.variable_symbol,
                expected_direction="credit",
                notes=f"{invoice.payment_scenario} for {invoice.invoice_no}",
            ))
            matched_ids.append(tx_id)
        elif invoice.payment_scenario == "unpaid":
            # No transaction at all — the month stays outstanding. Mirrors the
            # app's "Unpaid · NO_PAYMENT_FOUND" reminder-eligible row.
            expected_status = "Unpaid"
            expected_reason = "no_payment_found"
            expected_action = "send_reminder"
            expected_received = Decimal("0.00")
        else:
            raise ValueError(f"Unsupported payment scenario: {invoice.payment_scenario}")

        reconciliation_expectations.append(ReconciliationExpectation(
            customer_id=invoice.customer_id,
            invoice_no=invoice.invoice_no,
            variable_symbol=invoice.variable_symbol,
            expected_status=expected_status,
            expected_reason=expected_reason,
            expected_invoice_action=expected_action,
            expected_received_total=decimal_string(expected_received),
            expected_matched_transaction_count=str(len(matched_ids)),
            expected_matched_transaction_ids=",".join(matched_ids),
        ))

    credit_noise_total = noise_count(len(invoices))
    debit_noise_total = noise_count(len(invoices))

    for noise_index in range(credit_noise_total):
        pattern_kind, reference_text, variable_symbol = UNKNOWN_CREDIT_PATTERNS[noise_index % len(UNKNOWN_CREDIT_PATTERNS)]
        tx_id = build_transaction_id("TX", transaction_number)
        transaction_number += 1
        booking = invoices[0].issue_date + dt.timedelta(days=(noise_index % 7) + 1)
        transactions.append(transaction_row(
            import_batch_id=import_batch_id,
            transaction_id=tx_id,
            booking_date=booking,
            amount=money(30 + noise_index * 15),
            direction="credit",
            payer_name=f"Unknown Payer {noise_index + 1}",
            counterparty_iban="",
            reference_text=reference_text,
            variable_symbol=variable_symbol,
            raw_description=reference_text,
            raw_source_reference={
                "scenario": pattern_kind,
                "type": "unknown_credit",
            },
            imported_at=imported_at,
            constant_symbol="0308" if variable_symbol else "",
        ))
        tx_expectations.append(TransactionExpectation(
            transaction_id=tx_id,
            expected_category="unknown_credit",
            expected_customer_id="",
            expected_variable_symbol=variable_symbol,
            expected_direction="credit",
            notes=f"Noise credit {pattern_kind}",
        ))
        if variable_symbol and variable_symbol in real_variable_symbols:
            raise ValueError("Noise credit reused a real variable symbol")

    for noise_index in range(debit_noise_total):
        vendor_name, memo = VENDOR_DEBIT_PATTERNS[noise_index % len(VENDOR_DEBIT_PATTERNS)]
        tx_id = build_transaction_id("TX", transaction_number)
        transaction_number += 1
        booking = invoices[0].issue_date + dt.timedelta(days=(noise_index % 7) + 2)
        vendor_bank_id = f"{9900 + noise_index:04d}"
        vendor_account = f"{8000000000 + noise_index:010d}"[-10:]
        vendor_iban = f"SK00{vendor_bank_id}000000{vendor_account}"
        transactions.append(transaction_row(
            import_batch_id=import_batch_id,
            transaction_id=tx_id,
            booking_date=booking,
            amount=money(70 + noise_index * 23),
            direction="debit",
            payer_name=vendor_name,
            counterparty_iban=vendor_iban,
            reference_text=memo,
            variable_symbol="",
            raw_description=memo,
            raw_source_reference={
                "scenario": "vendor_debit",
                "vendor_name": vendor_name,
            },
            imported_at=imported_at,
        ))
        tx_expectations.append(TransactionExpectation(
            transaction_id=tx_id,
            expected_category="vendor_debit",
            expected_customer_id="",
            expected_variable_symbol="",
            expected_direction="debit",
            notes=f"Vendor debit {vendor_name}",
        ))

    transactions.sort(key=lambda transaction: (transaction.booking_date, transaction.transaction_id))
    tx_expectations.sort(key=lambda item: item.transaction_id)
    reconciliation_expectations.sort(key=lambda item: item.variable_symbol)

    min_booking = min(parse_date(transaction.booking_date) for transaction in transactions)
    max_booking = max(parse_date(transaction.booking_date) for transaction in transactions)
    summary = {
        "batch_id": invoices[0].batch_id,
        "import_batch_id": import_batch_id,
        "source": "generated_tatra_ofx",
        "invoice_count": len(invoices),
        "transaction_count": len(transactions),
        "transaction_expectation_count": len(tx_expectations),
        "reconciliation_expectation_count": len(reconciliation_expectations),
        "duplicate_exact_count": sum(1 for row in reconciliation_expectations if row.expected_reason == "duplicate_possible"),
        "noise_credit_count": credit_noise_total,
        "noise_debit_count": debit_noise_total,
        "supplier_name": invoices[0].supplier_name,
        "supplier_iban": invoices[0].supplier_iban,
        "statement_bank_id": supplier_bank_id,
        "statement_account_id": supplier_account_id,
        "currency": "EUR",
        "imported_at": imported_at,
        "date_start": min_booking.isoformat(),
        "date_end": max_booking.isoformat(),
        "seed": seed,
    }
    return transactions, tx_expectations, reconciliation_expectations, summary


def format_bank_date(value: str) -> str:
    return value.replace("-", "")


def write_ofx(path: Path, transactions: list[BankTransaction], summary: dict[str, Any]) -> None:
    root = ET.Element("OFX")
    statement = ET.SubElement(root, "STMTRS")
    account_from = ET.SubElement(statement, "BANKACCTFROM")
    ET.SubElement(account_from, "BANKID").text = str(summary["statement_bank_id"])
    ET.SubElement(account_from, "ACCTID").text = str(summary["statement_account_id"])
    ET.SubElement(account_from, "IBAN").text = str(summary["supplier_iban"])

    tx_list = ET.SubElement(statement, "BANKTRANLIST")
    ET.SubElement(tx_list, "DTSTART").text = format_bank_date(str(summary["date_start"]))
    ET.SubElement(tx_list, "DTEND").text = format_bank_date(str(summary["date_end"]))

    for transaction in transactions:
        entry = ET.SubElement(tx_list, "STMTTRN")
        ET.SubElement(entry, "TRNTYPE").text = "CREDIT" if transaction.direction == "credit" else "DEBIT"
        ET.SubElement(entry, "DTPOSTED").text = format_bank_date(transaction.booking_date)
        ET.SubElement(entry, "DTAVAIL").text = format_bank_date(transaction.value_date)
        ET.SubElement(entry, "TRNAMT").text = transaction.amount
        ET.SubElement(entry, "FITID").text = transaction.transaction_id
        if transaction.variable_symbol:
            ET.SubElement(entry, "TRNVASYM").text = transaction.variable_symbol
        if transaction.constant_symbol:
            ET.SubElement(entry, "TRNCOSYM").text = transaction.constant_symbol
        ET.SubElement(entry, "REFERENCE_E2E").text = transaction.reference_text
        ET.SubElement(entry, "NAME").text = transaction.payer_name
        counterparty = ET.SubElement(entry, "BANKACCTTO")
        iban = compact_iban(transaction.counterparty_iban)
        if re.fullmatch(r"SK\d{22}", iban):
            bank_id, account_id = parse_sk_iban(iban)
            ET.SubElement(counterparty, "BANKID").text = bank_id
            ET.SubElement(counterparty, "ACCTID").text = account_id
            ET.SubElement(counterparty, "IBAN").text = iban
        else:
            ET.SubElement(counterparty, "BANKID").text = ""
            ET.SubElement(counterparty, "ACCTID").text = ""
            ET.SubElement(counterparty, "IBAN").text = iban
        ET.SubElement(entry, "MEMO").text = transaction.raw_description
        ET.SubElement(entry, "CURRENCY").text = transaction.currency

    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    tree.write(path, encoding="utf-8", xml_declaration=True)


def default_outdir(invoices_path: Path, batch_id: str) -> Path:
    return invoices_path.parent.parent / f"bank_statement_{batch_id}"


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate deterministic Tatra-style bank statement data from invoices.csv.")
    parser.add_argument("--invoices", required=True, type=Path, help="Path to invoice manifests/invoices.csv.")
    parser.add_argument("--outdir", type=Path, help="Output directory. Defaults to a sibling bank_statement_<batch_id> directory.")
    parser.add_argument("--seed", default=42, type=int, help="Deterministic seed for booking-date and noise distribution.")
    args = parser.parse_args()

    invoices = read_invoice_rows(args.invoices)
    outdir = args.outdir or default_outdir(args.invoices, invoices[0].batch_id)
    outdir.mkdir(parents=True, exist_ok=True)

    transactions, tx_expectations, reconciliation_expectations, summary = generate_statement(invoices, seed=args.seed)

    write_csv(outdir / "transactions.csv", csv_rows(transactions))
    write_csv(outdir / "transaction_expectations.csv", csv_rows(tx_expectations))
    write_csv(outdir / "reconciliation_expectations.csv", csv_rows(reconciliation_expectations))
    write_json(outdir / "transaction_expectations.json", [asdict(item) for item in tx_expectations])
    write_json(outdir / "reconciliation_expectations.json", [asdict(item) for item in reconciliation_expectations])
    write_json(outdir / "summary.json", summary)
    write_ofx(outdir / "statement.ofx", transactions, summary)

    print(f"Generated {len(transactions)} transactions in: {outdir / 'transactions.csv'}")
    print(f"Generated OFX statement in: {outdir / 'statement.ofx'}")
    print(f"Generated expectations in: {outdir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
