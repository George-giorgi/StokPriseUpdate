# Test Invoice Launcher

This is a simple clickable prototype for the invoice workflow.

## Double-click to use it

- Double-click `start_test_invoice.cmd` in the `Test` folder.
- Choose the PDF folder you want to process.
- Choose or confirm the workbook path.
- The script will process the invoice PDFs and save the updated workbook.

## Default folders

- PDFs: `Test/Test_new_Invoices`
- Workbook output: `Test/Test_MATERIAL_PRICES.xlsx`

## Notes

This is a prototype for the same pattern that will be reused for each supplier parser.

For a cleaner Windows experience, `start_test_invoice.cmd` tries to launch the GUI through `pythonw` first so it feels like a normal desktop app instead of a command-line tool.
