# Bulk Zálohová Faktúra Generator

Instructions for generating many DOCX and PDF invoice/proforma-invoice documents from the prepared Word template.

The generator creates unique sample customers, fills the invoice template, adds a QR code, optionally renders text-based PDFs directly from Python, and writes invoice/customer manifests in CSV and JSON.

---

## 1. Requirements

Install Python 3.10 or newer.

Install the required Python packages:

```bash
pip install python-docx qrcode[pil] pillow
```

For PDF export, install the Python dependencies from `requirements.txt`. The current PDF path is rendered directly from Python and does not need LibreOffice.

---

## 2. Files Needed

Place these files in the same working folder:

```text
bulk_zalohova_faktura_generator.py
zalohova_faktura_template_ready.docx
```

Recommended folder layout:

```text
invoice-generator/
  bulk_zalohova_faktura_generator.py
  zalohova_faktura_template_ready.docx
  customers.csv              # optional
  generated_invoices/        # created by the script
```

---

## 3. Generate 1000 DOCX Documents With Fake Sample Customers

Run:

```bash
python bulk_zalohova_faktura_generator.py \
  --template zalohova_faktura_template_ready.docx \
  --outdir ./generated_invoices \
  --count 1000 \
  --issue-date 2026-05-25
```

This creates:

```text
generated_invoices/
  docx/
  manifests/
  _qr_tmp/
```

The `docx/` folder contains the generated Word documents.
The `manifests/` folder contains `invoices.csv`, `invoices.json`, `customers.csv`, and `customers.json`.

---

## 4. Generate DOCX and PDF Documents

Run:

```bash
python bulk_zalohova_faktura_generator.py \
  --template zalohova_faktura_template_ready.docx \
  --outdir ./generated_invoices \
  --count 1000 \
  --issue-date 2026-05-25 \
  --pdf
```

This creates:

```text
generated_invoices/
  docx/
  pdf/
  manifests/
  _qr_tmp/
```

The `pdf/` folder contains the converted PDF files.

---

## 5. Use a Real Customer CSV Instead of Fake Customers

Create a file named `customers.csv` with these headers:

```csv
customer_name,street,city_country,ico,dic,ic_dph,register,contact,phone,email,note,customer_id,customer_iban,customer_bank_code,customer_account_number,customer_bank_name
```

Example:

```csv
customer_name,street,city_country,ico,dic,ic_dph,register,contact,phone,email,note,customer_id,customer_iban,customer_bank_code,customer_account_number,customer_bank_name
Novák Consulting s.r.o.,Hlavná 12,040 01 Košice, Slovensko,12345678,2123456789,SK2123456789,"Zapísaná v OR OS Košice I, oddiel Sro, vložka 12345/V",Ján Novák,+421 900 123 456,jan.novak@example.sk,Testovací zákazník,CUST-CSV-001,SK1211000000007000000001,1100,7000000001,Tatra banka
Kováčová Design s.r.o.,Dunajská 8,811 08 Bratislava, Slovensko,87654321,2187654321,SK2187654321,"Zapísaná v OR OS Bratislava I, oddiel Sro, vložka 98765/B",Eva Kováčová,+421 911 222 333,eva.kovacova@example.sk,Ukážkový záznam,,,,,
```

Then run:

```bash
python bulk_zalohova_faktura_generator.py \
  --template zalohova_faktura_template_ready.docx \
  --outdir ./generated_invoices \
  --customers-csv customers.csv \
  --count 1000 \
  --issue-date 2026-05-25 \
  --pdf
```

If the CSV has fewer rows than `--count`, the script continues with generated sample customers for the remaining documents and fills in missing optional customer fields deterministically.

---

## 6. Common Arguments

| Argument | Example | Description |
|---|---:|---|
| `--template` | `zalohova_faktura_template_ready.docx` | Path to the DOCX template |
| `--outdir` | `./generated_invoices` | Output folder |
| `--count` | `1000` | Number of documents to generate |
| `--issue-date` | `2026-05-25` | Invoice issue date in `YYYY-MM-DD` format |
| `--customers-csv` | `customers.csv` | Optional customer data source |
| `--pdf` | enabled flag | Also export PDFs |
| `--soffice` | any value | Deprecated compatibility flag; ignored |

---

## 7. Output Naming

Generated files use invoice-like numbering, for example:

```text
ZF-20260525-0001.docx
ZF-20260525-0002.docx
ZF-20260525-0003.docx
```

When PDF export is enabled:

```text
ZF-20260525-0001.pdf
ZF-20260525-0002.pdf
ZF-20260525-0003.pdf
```

## 8. Manifest Fields

The invoice manifest includes fields needed for later Tatra import/export testing, including:

- `customer_id`
- `customer_name`
- `variable_symbol`
- `billing_month`
- `invoice_total_amount`
- `charge_type` (`monthly` or `annual_extra`)
- `amount_bucket`
- `payment_scenario`
- `simulated_paid_total`
- `simulated_split_amounts`
- `reference_text_template`

The expected-charges manifest carries `charge_amount` and `charge_type` per charge. The customer manifest includes one deduplicated row per customer with generated or normalized bank identifiers plus billing-profile columns (`monthly_fee`, `annual_extra_fee`, `annual_extra_interval_months`, `payment_method`, `vat_status`, `status`).

## 9. QR Code Warning

The generated QR code currently contains plain invoice/payment sample text.

It is **not** a proper Slovak bank payment QR code using official **PAY by square** encoding.

For real payment QR support, replace the function:

```python
def make_qr_payload(...):
```

with a proper PAY by square encoder.

The rest of the script is already structured so the QR payload can be swapped without rewriting the document generation logic.

## 10. Practical Testing Flow

Start with a small batch first:

```bash
python bulk_zalohova_faktura_generator.py \
  --template zalohova_faktura_template_ready.docx \
  --outdir ./test_invoices \
  --count 5 \
  --issue-date 2026-05-25 \
  --pdf
```

Check:

1. The DOCX opens correctly in Word or Google Docs.
2. The layout still fits on one page.
3. The QR code appears in the correct area.
4. The total amount and written amount match.
5. The generated PDF contains readable text and the expected totals.

Only after that, run the full 1000-document batch.

## 11. Performance Notes

Generating DOCX files is usually fast.

Direct PDF rendering is usually faster and simpler than office-suite conversion, but 1000 files can still take a while depending on CPU and disk speed.

Recommended workflow:

1. Generate DOCX files first.
2. Inspect a few samples.
3. Run PDF export after confirming the layout.
4. Keep the generated DOCX files as editable source files.

---

## 12. Safety Notes

For real customer data:

- Do not commit generated invoices to Git.
- Do not leave customer CSV files in public folders.
- Do not upload real invoice batches to random online converters.
- Keep sample/test data separate from production data.
- Use realistic but fake test data when testing layouts.

Because leaking 1000 invoices would be a memorable way to make accounting exciting, and not in the good way.
