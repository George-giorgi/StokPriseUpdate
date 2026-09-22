# Southern Invoice Price Update

This workflow processes Southern Building Products invoices separately from the other suppliers.

## Folders and files

- `Southeren/Southeren_new_Invoices/` contains new Southern PDF invoices ready for processing.
- `Southeren/Southeren_full_Invoices/` stores older or processed Southern invoices.
- `Southeren/Southern_MATERIAL_PRICES.xlsx` is the separate Southern workbook.
- `Southeren/southern_invoice_import.py` is the Southern-specific parser.

The existing folder is named `Southeren`, while the supplier is Southern Building Products. The parser uses the printed invoice Code as the Excel item code and joins wrapped detail lines.

## Run command

Run from the project folder:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
python Southeren\southern_invoice_import.py "Southeren/Southeren_new_Invoices" "Southeren/Southern_MATERIAL_PRICES.xlsx"
```

If the workbook does not exist, the script creates a clean Southern workbook automatically.

After processing, check the workbook backup and CSV report, then move processed PDFs to `Southeren_full_Invoices`.
