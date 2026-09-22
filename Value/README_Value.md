# Value Invoice Price Update

This guide explains how to process Value supplier invoices and update the separate Value Excel workbook.

## Folders and files

### `Value/Value_new_Invoices/`

Put new, ready-to-process Value invoice PDFs in this folder.

Example:

```text
Value_new_Invoices/
    Invoice_308293.pdf
    Invoice_310189.pdf
```

Only PDF files in this folder are processed.

### `Value/Value_MATERIAL_PRICES.xlsx`

This is the separate Value workbook. It stores Value materials, prices, invoice history, and update results.

Do not use the Chadwicks workbook for Value invoices.

### `value_invoice_import.py`

This is the Value-specific parser and updater. It uses the Value invoice layout and updates only the Value workbook.

## How to run

1. Open Command Prompt or PowerShell.
2. Go to the project folder:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
```

3. Put the new Value PDFs into `Value/Value_new_Invoices`.
4. Run:

```cmd
python value_invoice_import.py "Value/Value_new_Invoices" "Value/Value_MATERIAL_PRICES.xlsx"
```

If the Value workbook does not exist, the script creates an empty Value workbook automatically.

## What the script does

The script:

- Reads Value invoice PDFs from `Value_new_Invoices`.
- Extracts the Value item code, description, unit, quantity, and price.
- Adds new Value materials to the workbook.
- Updates changed Value prices.
- Keeps previous prices and price history.
- Creates a backup workbook before updating.
- Creates a CSV report after processing.
- Sends the configured email report when email settings are available.

## After running

Check the command window for:

```text
Found ... PDF invoice(s).
Items detected: ...
Prices updated: ...
Prices unchanged: ...
New materials: ...
Errors: ...
```

Then check:

- `Value/Value_MATERIAL_PRICES.xlsx`
- The generated backup workbook in the `Value` folder
- The generated CSV report in the `Value` folder

## Next batch

After confirming the update is correct, move or remove the processed PDFs from `Value_new_Invoices` before adding the next batch. This prevents the same invoices from being processed again.

## Important

- Keep Value invoices and the Value workbook separate from Chadwicks and Archers files.
- Do not delete `Value_MATERIAL_PRICES.xlsx`; it contains the accumulated Value price history.
- Do not store Gmail passwords in the script.
- Keep real invoices private and out of GitHub.
