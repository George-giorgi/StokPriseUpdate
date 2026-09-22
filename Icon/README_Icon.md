# Icon Invoice Price Update

This workflow processes Icon Building Products invoices separately from the other suppliers.

## Folders and files

- `Icon/Icon_new_Invoices/` contains new Icon PDF invoices ready for processing.
- `Icon/Icon_full_Invoices/` stores older or processed Icon invoices.
- `Icon/Icon_MATERIAL_PRICES.xlsx` is the separate Icon workbook.
- `Icon/icon_invoice_import.py` is the Icon-specific parser.

The parser uses the printed Icon item code as the Excel item code, extracts quantity and unit price, and joins wrapped description lines. Duplicate copies of the same invoice are deduplicated during parsing.

## Run command

Run from the project folder:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
python Icon\icon_invoice_import.py "Icon/Icon_new_Invoices" "Icon/Icon_MATERIAL_PRICES.xlsx"
```

If the workbook does not exist, the script creates a clean Icon workbook automatically.

After processing, check the workbook backup and CSV report, then move processed PDFs to `Icon_full_Invoices`.
