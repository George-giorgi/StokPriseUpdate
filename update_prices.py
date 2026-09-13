"""
TEST VERSION — updates MATERIAL_PRICES from invoice data exported to CSV.

CSV columns:
Invoice No.,Supplier,Invoice Date,Material ID,Material,Quantity,Unit Price

Logic:
Previous Price <- current New Price
New Price      <- latest invoice Unit Price

Run:
python update_prices.py invoice_items.csv WarehouseFlow_EWI_With_Prices_And_Invoices.xlsx
"""
import csv, sys
from openpyxl import load_workbook

if len(sys.argv) != 3:
    print("Usage: python update_prices.py invoice_items.csv workbook.xlsx")
    raise SystemExit(1)

csv_file, workbook_file = sys.argv[1], sys.argv[2]
wb = load_workbook(workbook_file)
ws = wb["MATERIAL_PRICES"]

header = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}

required = ["Material ID", "New Price (€)", "Previous Price (€)", "Last Invoice", "Last Updated"]
missing = [x for x in required if x not in header]
if missing:
    raise ValueError("Missing columns: " + ", ".join(missing))

rows_by_id = {}
for r in range(2, ws.max_row + 1):
    mid = ws.cell(r, header["Material ID"]).value
    if mid is not None:
        rows_by_id[str(mid).strip()] = r

updated = 0
unknown = []

with open(csv_file, newline="", encoding="utf-8-sig") as f:
    for item in csv.DictReader(f):
        mid = str(item.get("Material ID", "")).strip()
        price = float(str(item["Unit Price"]).replace("€","").replace(",","").strip())
        invoice = item.get("Invoice No.", "")
        date = item.get("Invoice Date", "")
        if mid not in rows_by_id:
            unknown.append(mid)
            continue

        r = rows_by_id[mid]
        old_new = ws.cell(r, header["New Price (€)"]).value
        ws.cell(r, header["Previous Price (€)"]).value = old_new
        ws.cell(r, header["New Price (€)"]).value = price
        ws.cell(r, header["Last Invoice"]).value = invoice
        ws.cell(r, header["Last Updated"]).value = date
        updated += 1

wb.save(workbook_file)
print(f"Updated prices: {updated}")
if unknown:
    print("Unknown Material IDs (review manually):")
    for x in sorted(set(unknown)):
        print(" -", x)
