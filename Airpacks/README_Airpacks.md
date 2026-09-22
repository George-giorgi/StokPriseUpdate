# Airpacks Invoice Price Update

## Folders and files

- `Airpacks/Airpacks_new_Invoices/` contains ready Airpacks PDF invoices.
- `Airpacks/Airpacks_MATERIAL_PRICES.xlsx` is the separate Airpacks workbook.
- `Airpacks/airpacks_invoice_import.py` is the Airpacks-specific parser.

## Run from the project folder

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
python Airpacks\airpacks_invoice_import.py "Airpacks/Airpacks_new_Invoices" "Airpacks/Airpacks_MATERIAL_PRICES.xlsx"
```

If the workbook does not exist, the script creates an empty Airpacks workbook automatically.

The parser updates only the Airpacks workbook. Keep Airpacks invoices separate from Chadwicks, Archers, and Value invoices.

Do not place invoices from another supplier in this folder. For example, `SI-061859.pdf` is a Molloy Concrete invoice and must be moved to a separate Molloy invoice workflow.

After processing, review the workbook backup and CSV report in the `Airpacks` folder before moving processed PDFs out of `Airpacks_new_Invoices`.
