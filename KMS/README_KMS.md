# KMS Invoice Price Update

This workflow processes KMS Concepts invoices separately from the other suppliers.

## Folders and files

- `KMS/KMS_new_Invoices/` contains new KMS PDF invoices ready for processing.
- `KMS/KMS_full_Invoices/` stores older or processed KMS invoices.
- `KMS/KMS_MATERIAL_PRICES.xlsx` is the separate KMS workbook.
- `KMS/kms_invoice_import.py` is the KMS-specific parser.

The parser uses the printed KMS invoice Code as the Excel item code, extracts quantity and price, and ignores carriage/delivery charges.

## Run command

Run from the project folder:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
python KMS\kms_invoice_import.py "KMS/KMS_new_Invoices" "KMS/KMS_MATERIAL_PRICES.xlsx"
```

If the workbook does not exist, the script creates a clean KMS workbook automatically.

After processing, check the workbook backup and CSV report, then move processed PDFs to `KMS_full_Invoices`.
