# Invoice Generator

Bulk generator for Slovak `Zálohová faktúra` documents from the prepared DOCX template.

The project currently generates:

- DOCX invoices from [`docs/zalohova_faktura_template_ready.docx`](/home/disane/Development/AccountingAutomation/invoice_generator/docs/zalohova_faktura_template_ready.docx)
- Optional text-based PDF exports through a direct Python renderer
- QR code images embedded into each generated document
- Invoice manifests in `CSV` and `JSON`
- Customer manifests in `CSV` and `JSON`
- Deterministic Tatra-style bank statement test data from `invoices.csv`

## Requirements

- Python 3.10+

Python 3.12.3 has been set up locally for this workspace.

## Setup

Create and activate the virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Installed Python Packages

The current environment uses:

- `python-docx==1.2.0`
- `qrcode[pil]==8.2`
- `pillow==12.2.0`
- `reportlab==4.5.1`

## Quick Start

Generate 1000 DOCX invoices:

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir generated_invoices \
  --count 1000 \
  --issue-date 2026-05-25
```

This now also writes:

```text
generated_invoices/
  manifests/
    invoices.csv
    invoices.json
    customers.csv
    customers.json
```

Generate 1000 DOCX + PDF invoices:

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir generated_invoices \
  --count 1000 \
  --issue-date 2026-05-25 \
  --pdf
```

Optional compatibility note:

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir generated_invoices \
  --count 1000 \
  --issue-date 2026-05-25 \
  --pdf \
  --soffice /path/to/soffice
```

`--soffice` is now ignored. It remains accepted only so older command examples do not break.

Generate a Tatra-style statement from an invoice manifest:

```bash
.venv/bin/python scripts/generate_tatra_bank_statement.py \
  --invoices generated_invoices/manifests/invoices.csv \
  --outdir generated_invoices/bank_statement \
  --seed 42
```

## Customer CSV

Use `--customers-csv` to supply real customer records instead of generated sample customers.

Expected headers:

```csv
customer_name,street,city_country,ico,dic,ic_dph,register,contact,phone,email,note,customer_id,customer_iban,customer_bank_code,customer_account_number,customer_bank_name
```

Behavior:

- If the CSV contains fewer than `--count` rows, the remainder is backfilled with synthetic customers.
- Missing optional fields in CSV rows are enriched deterministically.
- CSV rows are preserved first; synthetic rows fill the rest.

## Amounts And Scenarios

For a 1000-invoice batch, the generator uses this deterministic mix:

- `300` invoices at `80.00 EUR`
- `300` invoices at `180.00 EUR`
- `250` invoices at `210.00 EUR`
- `150` invoices at random whole-EUR amounts in `1..9999`, excluding `80`, `180`, and `210`

Every invoice manifest row also includes reconciliation-test metadata:

- `amount_bucket`
- `payment_scenario`
- `suggested_received_amount`
- `suggested_split_amounts`
- `reference_text_template`

Supported payment scenarios:

- `exact_single`
- `exact_split_total`
- `underpay`
- `overpay`

## Repo Layout

```text
invoice_generator/
  .venv/
  docs/
    zalohova_faktura_sample_filled_words_only.docx
    zalohova_faktura_sample_filled_words_only.pdf
    zalohova_faktura_template_ready.docx
  scripts/
    bulk_zalohova_faktura_generator.py
    bulk_zalohova_faktura_generator_instructions.md
  tests/
    test_bulk_zalohova_faktura_generator.py
  requirements.txt
  README.md
  AGENT.md
  CLAUDE.md
```

## Notes

- `--count` defaults to `1000`, so large-batch generation is already the default path.
- `--pdf` now generates parsable text PDFs directly from Python without LibreOffice.
- The QR payload is test text, not a Slovak `PAY by square` implementation.
- `invoices.csv` is the primary downstream flat export for later Tatra transaction generation.
- `generate_tatra_bank_statement.py` reads `invoices.csv` and emits `transactions.csv`, `statement.ofx`, expectations files, and `summary.json`.

## Verification

Run the unit tests:

```bash
.venv/bin/python -m unittest discover -s tests -v
```
# invoice_and_bank_statement_generator
