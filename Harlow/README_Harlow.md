# Harlow Invoice Price Update

This workflow processes Harlow Agencies invoices separately from the other suppliers.

## Folders and files

- `Harlow/Harlow_new_Invoices/` contains new Harlow PDF invoices ready for processing.
- `Harlow/Harlow_full_Invoices/` stores older or processed Harlow invoices.
- `Harlow/Harlow_MATERIAL_PRICES.xlsx` is the separate Harlow workbook.
- `Harlow/harlow_invoice_import.py` is the Harlow-specific parser.

The parser reads the Harlow invoice header, extracts the invoice number and date, and reads the table rows using the printed item code, quantity, description and unit price.

## Run command

Run from the project folder:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
python Harlow\harlow_invoice_import.py "Harlow/Harlow_new_Invoices" "Harlow/Harlow_MATERIAL_PRICES.xlsx"
```

If the workbook does not exist, the script creates a clean Harlow workbook automatically.

After processing, check the workbook backup and CSV report, then move processed PDFs to `Harlow/Harlow_full_Invoices`.
