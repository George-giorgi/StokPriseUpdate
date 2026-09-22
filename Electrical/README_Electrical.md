# Electrical Invoice Price Update

This workflow processes Kellihers Electrical invoices separately from the other suppliers.

## Folders and files

- `Electrical/Electrical_new_Invoices/` contains new Electrical PDF invoices ready for processing.
- `Electrical/Electrical_full_Invoices/` stores older or processed Electrical invoices.
- `Electrical/Electrical_MATERIAL_PRICES.xlsx` is the separate Electrical workbook.
- `Electrical/electrical_invoice_import.py` is the Electrical-specific parser.

The parser uses the invoice Item Number as the Excel item code and includes the Config Number in the description.

## Run command

Run from the project folder:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
python Electrical\electrical_invoice_import.py "Electrical/Electrical_new_Invoices" "Electrical/Electrical_MATERIAL_PRICES.xlsx"
```

If the workbook does not exist, the script creates a clean Electrical workbook automatically.

## Workflow

1. Put new Electrical PDFs in `Electrical_new_Invoices`.
2. Run the command above.
3. Check the updated Electrical workbook, backup, and CSV report.
4. Move processed PDFs to `Electrical_full_Invoices` before the next batch.

Keep Electrical invoices and the Electrical workbook separate from Chadwicks, Archers, Value, and Airpacks files.
