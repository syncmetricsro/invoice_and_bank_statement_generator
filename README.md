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

To set up the project locally after cloning it from GitHub:

```bash
git clone <repository-url>
cd invoice_generator
```

Create and activate the virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Verify the installation by running the test suite:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

All generated output folders (`generated_invoices*/`, including the reference batches `generated_invoices_5/` and `generated_invoices_25/`) are gitignored because they are fully deterministic and can be regenerated at any time — after a fresh clone, recreate them with the commands in [Manual Test Scenarios](#manual-test-scenarios).

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
    expected_charges.csv
    expected_charges.json
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

## Manual Test Scenarios

Two reference batches are used for manual reconciliation testing. Both are fully deterministic (seed `42`, issue date `2026-06-10`), so regenerating them always reproduces the same customers, variable symbols, and amounts.

### `generated_invoices_5` — smallest end-to-end batch

5 customers, 6 invoices (CUST-00005 also gets the annual extra). Covers one of each basic scenario: exact match, split payment (90+90), underpayment, overpayment, and a paid annual-extra invoice. No duplicate case — duplicates only appear on every 10th exact payment.

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir generated_invoices_5 \
  --count 5 \
  --issue-date 2026-06-10 \
  --pdf

.venv/bin/python scripts/generate_tatra_bank_statement.py \
  --invoices generated_invoices_5/manifests/invoices.csv \
  --outdir generated_invoices_5/bank_statement \
  --seed 42
```

### `generated_invoices_25` — smallest batch with a duplicate payment

25 customers, 30 invoices (annual extras for customers 5, 10, 15, 20, 25), 40 transactions. Reconciliation mix: 14 exact matches, 5 splits, 5 underpayments, 5 overpayments, and exactly 1 duplicate — CUST-00025 pays invoice `2026-0025` (€80, VS `20260025`) twice on consecutive days, with two different reference-text formats, so matching must rely on the variable symbol.

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir generated_invoices_25 \
  --count 25 \
  --issue-date 2026-06-10 \
  --pdf

.venv/bin/python scripts/generate_tatra_bank_statement.py \
  --invoices generated_invoices_25/manifests/invoices.csv \
  --outdir generated_invoices_25/bank_statement \
  --seed 42
```

Rule of thumb for sizing your own batch: about 60% of invoices pay exactly (40% of monthlies plus all annual extras), and every 10th exact payment becomes a duplicate — so a batch of `N` yields roughly `0.6N / 10` duplicate cases (a 1000-batch yields 60).

## Customer CSV

Use `--customers-csv` to supply real customer records instead of generated sample customers.

Expected headers:

```csv
customer_name,street,city_country,ico,dic,ic_dph,register,contact,phone,email,note,customer_id,customer_iban,customer_bank_code,customer_account_number,customer_bank_name
```

Optional billing-profile overrides are also accepted: `payment_method`, `vat_status`, `status`, `annual_extra_fee`, `annual_extra_interval_months`. When empty, they are derived deterministically. `monthly_fee` is never read from the CSV — it always equals the assigned invoice amount.

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

Every 5th customer additionally receives an annual extra invoice (`120.00 EUR`, `exact_single`), so `--count N` produces `N + floor(N/5)` documents. Annual extra invoices use invoice numbers ending in `-AE` and 9-digit variable symbols, so they never collide with the 8-digit monthly symbols.

Every invoice manifest row also includes reconciliation-test metadata:

- `charge_type` (`monthly` or `annual_extra`)
- `amount_bucket`
- `payment_scenario`
- `simulated_paid_total` (the amount the generated bank statement will simulate paying)
- `simulated_split_amounts`
- `reference_text_template`

The generator also writes a dedicated expected-charge export for downstream seeding:

- `expected_charges.csv`
- `expected_charges.json`

Each expected-charge row carries the owed amount and payment deadline:

- `customer_id`
- `customer_name`
- `billing_month`
- `charge_type`
- `variable_symbol`
- `charge_amount`
- `due_date`
- `invoice_no`

The customers manifest carries a billing profile per customer: `monthly_fee` (equals the assigned invoice amount), `annual_extra_fee` and `annual_extra_interval_months` (set for every 5th customer), `payment_method` (`bank_transfer`, every 10th `direct_debit`), `vat_status` (`vat_payer` / `non_vat_payer`, derived from `ic_dph`), and `status` (`active`, every 40th `inactive`).

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
