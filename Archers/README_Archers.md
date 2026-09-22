# Archers Invoice Price Update

## Folders and files

- `Archers/Archers_Invoices/` contains ready Archers PDF invoices.
- `Archers/Archers_MATERIAL_PRICES.xlsx` is the separate Archers workbook.
- `Archers/archers_invoice_import.py` is the Archers-specific parser.

## Run from the project folder

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
python Archers\archers_invoice_import.py "Archers/Archers_Invoices" "Archers/Archers_MATERIAL_PRICES.xlsx"
```

The parser updates only the Archers workbook. Keep Archers invoices separate from Chadwicks and Value invoices.

After processing, review the workbook backup and CSV report in the `Archers` folder before moving processed PDFs out of `Archers_Invoices`.
