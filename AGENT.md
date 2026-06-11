# AGENT.md

## Purpose

This repo contains a single Python batch document generator for Slovak advance invoices (`Zálohová faktúra`).

Primary script:

- [scripts/bulk_zalohova_faktura_generator.py](/home/disane/Development/AccountingAutomation/invoice_generator/scripts/bulk_zalohova_faktura_generator.py)
- [scripts/generate_tatra_bank_statement.py](/home/disane/Development/AccountingAutomation/invoice_generator/scripts/generate_tatra_bank_statement.py)

Primary template:

- [docs/zalohova_faktura_template_ready.docx](/home/disane/Development/AccountingAutomation/invoice_generator/docs/zalohova_faktura_template_ready.docx)

## Environment

- Use the local virtualenv at `.venv`
- Install deps with `pip install -r requirements.txt`
- Run commands with `.venv/bin/python`

## Current Dependency Set

- `python-docx`
- `qrcode[pil]`
- `pillow`
- `reportlab`

## Working Rules

- Preserve the current DOCX template path unless the user asks to replace it.
- Treat `docs/` assets as source artifacts, not generated output.
- Validate behavior with a small batch before assuming a 1000-document run is safe.
- Keep README examples aligned with the actual CLI behavior in the script.
- Keep manifest field names aligned with the downstream accounting automation flow, especially `customer_id`, `billing_month`, `variable_symbol`, `invoice_total_amount`, and `charge_amount`.
- Keep the bank-statement transaction fields aligned with the downstream normalized import shape, especially `transaction_id`, `direction`, `reference_text`, `variable_symbol`, and `counterparty_iban`.

## Important Behavioral Notes

- `--count` defaults to `1000`.
- If `--customers-csv` is shorter than `--count`, the script backfills the remainder with deterministic synthetic customers.
- PDF export is generated directly from Python with `reportlab`.
- QR codes are placeholder payment text, not official `PAY by square`.
- Every run writes `manifests/invoices.csv`, `manifests/invoices.json`, `manifests/customers.csv`, and `manifests/customers.json`.
- For a 1000-row batch, the amount distribution is `300x80`, `300x180`, `250x210`, and `150` random whole-EUR amounts.
- Payment scenarios are encoded into manifests as `exact_single`, `exact_split_total`, `underpay`, and `overpay`.
- `generate_tatra_bank_statement.py` reads `manifests/invoices.csv` and writes `transactions.csv`, `statement.ofx`, expectations files, and `summary.json`.
- The bank generator injects deterministic duplicate-payment and noise transactions to exercise reconciliation edge cases.

## Useful Commands

Show CLI help:

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py --help
```

Smoke test with one document:

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir tmp_smoke \
  --count 1 \
  --issue-date 2026-05-25
```

Run the unit tests:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Generate a bank statement from an invoice manifest:

```bash
.venv/bin/python scripts/generate_tatra_bank_statement.py \
  --invoices generated_invoices/manifests/invoices.csv \
  --outdir generated_invoices/bank_statement \
  --seed 42
```
