# CLAUDE.md

## Project Summary

This is a Python invoice batch generator centered around a DOCX template and optional direct Python PDF generation.

## Start Here

1. Activate the environment: `source .venv/bin/activate`
2. Main script: [scripts/bulk_zalohova_faktura_generator.py](/home/disane/Development/AccountingAutomation/invoice_generator/scripts/bulk_zalohova_faktura_generator.py)
3. Bank statement script: [scripts/generate_tatra_bank_statement.py](/home/disane/Development/AccountingAutomation/invoice_generator/scripts/generate_tatra_bank_statement.py)
4. Template: [docs/zalohova_faktura_template_ready.docx](/home/disane/Development/AccountingAutomation/invoice_generator/docs/zalohova_faktura_template_ready.docx)
5. Install packages with: `pip install -r requirements.txt`

## Commands

Run a standard 1000-document batch:

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir generated_invoices \
  --count 1000 \
  --issue-date 2026-05-25
```

Run the verification suite:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Run with PDF conversion:

```bash
.venv/bin/python scripts/bulk_zalohova_faktura_generator.py \
  --template docs/zalohova_faktura_template_ready.docx \
  --outdir generated_invoices \
  --count 1000 \
  --issue-date 2026-05-25 \
  --pdf
```

Recreate the manual-test reference batches `generated_invoices_5` (one of each basic scenario) and `generated_invoices_25` (adds one duplicate-payment case): see "Manual Test Scenarios" in [README.md](/home/disane/Development/AccountingAutomation/invoice_generator/README.md) — both use `--issue-date 2026-06-10`, seed `42`, `--pdf`, and a statement run into `<outdir>/bank_statement`.

Generate a Tatra-style statement from an invoice manifest:

```bash
.venv/bin/python scripts/generate_tatra_bank_statement.py \
  --invoices generated_invoices/manifests/invoices.csv \
  --outdir generated_invoices/bank_statement \
  --seed 42
```

## Implementation Notes

- The generator creates synthetic customer data when `--customers-csv` is omitted.
- If the customer CSV is shorter than `--count`, the remaining customers are synthesized automatically.
- `safe_filename()` limits filenames to 120 characters.
- `--pdf` renders text-based PDFs directly from Python with `reportlab`.
- Progress output is printed every 50 generated DOCX files.
- Sidecar manifests are always written to `outdir/manifests/`.
- For a 1000-row batch, amount buckets are `80`, `180`, `210`, and random whole-EUR values.
- Manifest rows include reconciliation metadata such as `payment_scenario`, `simulated_paid_total`, and `simulated_split_amounts`. Naming convention: `invoice_*`/`charge_*` = owed amounts, `simulated_*` = what the fake bank statement will do, `expected_*` = reconciliation ground truth.
- Invoice and expected-charge rows carry a `charge_type` (`monthly` or `annual_extra`). Every 5th customer gets an extra annual invoice (`120.00`/12m, `exact_single`, invoice number suffix `-AE`, 9-digit variable symbol), so `--count N` yields `N + floor(N/5)` documents.
- `customers.csv` includes billing-profile columns: `monthly_fee` (equals the assigned invoice amount), `annual_extra_fee`, `annual_extra_interval_months`, `payment_method`, `vat_status`, `status`.
- `generate_tatra_bank_statement.py` reads `manifests/invoices.csv` and writes `transactions.csv`, `statement.ofx`, expectations files, and `summary.json`.
- The statement generator adds deterministic duplicates and controlled noise credits/debits for reconciliation testing.

## Constraints

- Do not imply the QR data is production banking format.
- Prefer testing with `--count 1` or `--count 5` before a 1000-file run.
