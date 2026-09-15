import os
import re
import sys
import shutil
import csv
from datetime import datetime

import pdfplumber
from openpyxl import load_workbook


# ============================================================
# CHADWICKS INVOICE IMPORTER - FINAL FIXED V2
#
# Usage:
#   python chadwicks_invoice_import_FINAL_FIXED_V2.py "CHAD_Invoices" "Chadwicks_Invoice_Material_Master.xlsx"
#
# Fixes:
# - Recognises BOX, BOX50, BOX100, BOX200, etc.
# - Uses item code + quantity + unit as the row anchor.
# - Keeps multi-line descriptions together.
# - Stops descriptions at invoice/footer/company notices.
# - Does not treat carriage/delivery charge as a material price unless
#   it already exists in the master.
# - Matches by Supplier + Material ID.
# - Updates New Price / Previous Price.
# - Writes PRICE_HISTORY.
# - Creates a backup and CSV report.
# ============================================================


SUPPLIER = "Chadwicks"

# Common units found on Chadwicks invoices.
# BOX\d* deliberately accepts BOX100 / BOX50 / BOX200 / etc.
UNIT_RE = re.compile(
    r"^(?:"
    r"EACH|EA|"
    r"BOX\d*|"
    r"BALE|"
    r"BAG|"
    r"SHEET|"
    r"LENGTH|"
    r"SET|"
    r"ROLL|"
    r"TUBE|"
    r"PACK|"
    r"PAIR|"
    r"METRE|"
    r"M|"
    r"UNIT|"
    r"PALLET"
    r")$",
    re.I,
)

ITEM_CODE_RE = re.compile(
    r"^(?P<code>[A-Z]?\d{3,}(?:/[A-Z0-9]+)?)$",
    re.I,
)

PRICE_RE = re.compile(r"(?<!\d)(\d{1,6}[.,]\d{2})(?!\d)")

STOP_PHRASES = [
    "value ex. vat",
    "value ex vat",
    "value ex.",
    "value ex",
    "vat amount",
    "total euro",
    "total excl. vat",
    "total incl. vat",
    "total inc. vat",
    "total including vat",
    "invoice total",
    "goods collected",
    "we are moving to our new",
    "airways industrial estate",
    "on feb ",
    "chadwicks",
    "registered office",
    "company registration",
    "vat registration",
    "thank you",
]

HEADER_WORDS = {
    "item",
    "code",
    "description",
    "qty",
    "quantity",
    "unit",
    "price",
    "value",
    "discount",
    "vat",
}


def clean_text(value):
    if value is None:
        return ""
    value = str(value).replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value).strip()
    return value


def normalise_code(value):
    return clean_text(value).upper()


def to_price(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)

    s = clean_text(value).replace("€", "").replace(" ", "")
    if not s:
        return None

    # Irish/European PDFs can use comma decimals.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")

    try:
        return round(float(s), 2)
    except ValueError:
        return None


def is_stop_line(line):
    low = clean_text(line).lower()
    if not low:
        return False

    return any(phrase in low for phrase in STOP_PHRASES)


def is_item_code(line):
    line = clean_text(line)
    if not line:
        return False

    # Allow codes such as 13045 and D029/111.
    return bool(ITEM_CODE_RE.fullmatch(line))


def is_unit(line):
    line = clean_text(line)
    return bool(UNIT_RE.fullmatch(line))


def extract_prices(text):
    values = []
    for match in PRICE_RE.finditer(text):
        p = to_price(match.group(1))
        if p is not None:
            values.append(p)
    return values


def find_pdf_text(pdf_path):
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            if text:
                pages.append(text)
    return "\n".join(pages)


def extract_invoice_number(text, filename):
    patterns = [
        r"\bInvoice\s*(?:No\.?|Number)?\s*[:#]?\s*([A-Z]{1,5}\d{4,})\b",
        r"\b(SX\d{6,}|SN\d{5,})\b",
    ]

    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            return m.group(1).upper()

    m = re.search(r"\b(SX\d{6,}|SN\d{5,})\b", filename, re.I)
    return m.group(1).upper() if m else os.path.splitext(filename)[0]


def extract_invoice_date(text):
    patterns = [
        r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b",
        r"\b(\d{1,2}-\d{1,2}-\d{2,4})\b",
    ]

    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            return m.group(1)

    return ""


def clean_description_suffix(description, unit_price, qty):
    """Remove trailing invoice table numeric columns from a description."""
    description = clean_text(description)
    if not description:
        return description

    price = to_price(unit_price)
    quantity = to_price(qty)
    if price is None or quantity is None:
        return description

    def money(v):
        return rf"{v:.2f}".replace('.', r'[.,]')

    p = money(price)
    q = str(int(quantity)) if quantity.is_integer() else money(quantity)

    # Chadwicks extraction commonly produces:
    # DESCRIPTION UNIT_PRICE LINE_EX_VAT TOTAL_INC_VAT QTY
    patterns = [
        rf"\s+{p}\s+\d{{1,8}}[.,]\d{{2}}\s+\d{{1,8}}[.,]\d{{2}}\s+{q}$",
        rf"\s+{p}\s+\d{{1,8}}[.,]\d{{2}}\s+{q}$",
        rf"\s+{p}\s+{q}$",
    ]
    for pattern in patterns:
        cleaned = re.sub(pattern, '', description, flags=re.I).strip()
        if cleaned != description:
            return cleaned

    # Last-resort: if the description ends with QTY and contains the
    # exact unit price immediately before the numeric table columns,
    # remove all trailing numeric monetary fields plus quantity.
    tokens = description.split()
    while tokens and re.fullmatch(r'\d+(?:[.,]\d{2})?', tokens[-1]):
        # Do not strip legitimate product dimensions such as 100, 1200, etc.
        if tokens[-1].replace(',', '.') == q:
            tokens.pop()
        else:
            break
    while len(tokens) >= 2 and re.fullmatch(r'\d+[.,]\d{2}', tokens[-1]):
        tokens.pop()
        # Stop after removing a maximum of three monetary columns.
        if len(description.split()) - len(tokens) >= 4:
            break
    return ' '.join(tokens).strip()

def parse_rows_from_lines(text):
    """Parse the actual Chadwicks extracted table structure.

    Chadwicks text extraction places the numeric table columns on the same
    line as the item in most invoices, e.g.:

        CODE QTY UNIT DESCRIPTION UNIT_PRICE LINE_EX_VAT TOTAL_INC_VAT VAT_CODE

    Continuation description lines may follow the row.  The key rule is:
    **split the first item line at the numeric price columns before joining
    continuation lines**.  This prevents prices, line totals and VAT values
    becoming part of the material description.
    """
    raw_lines = [clean_text(x) for x in text.splitlines()]
    raw_lines = [x for x in raw_lines if x]

    # Actual Chadwicks item-row shape observed in the supplied invoices.
    inline_re = re.compile(
        r"^(?P<code>[A-Z]?\d{3,}(?:/[A-Z0-9]+)?)\s+"
        r"(?P<qty>\d+(?:[.,]\d+)?)\s+"
        r"(?P<unit>BOX\d*|EACH|EA|BALE|BAG|SHEET|LENGTH|SET|ROLL|TUBE|PACK|PAIR|METRE|M|UNIT|PALLET)\s+"
        r"(?P<description>.*?)\s+"
        r"(?P<unit_price>\d{1,8}[.,]\d{2})\s+"
        r"(?P<line_total>\d{1,10}[.,]\d{2})\s+"
        r"(?P<inc_total>\d{1,10}[.,]\d{2})\s+"
        r"(?P<vat_code>\d+)\s*$",
        re.I,
    )

    rows = []
    current = None

    def finish_current():
        nonlocal current
        if not current:
            return
        current["description"] = clean_text(current["description"])
        for marker in [
            "we are moving to our new",
            "airways industrial estate",
            "on feb ",
        ]:
            pos = current["description"].lower().find(marker)
            if pos >= 0:
                current["description"] = current["description"][:pos].strip()
        if current["description"] and current["unit_price"] is not None:
            rows.append(current)
        current = None

    for line in raw_lines:
        m = inline_re.match(line)
        if m:
            finish_current()
            current = {
                "code": normalise_code(m.group("code")),
                "qty": to_price(m.group("qty")),
                "unit": m.group("unit").upper(),
                "description": clean_text(m.group("description")),
                "unit_price": to_price(m.group("unit_price")),
            }
            continue

        # A new item code that did not match the complete inline row ends
        # the previous item; do not guess a price from unrelated footer data.
        if is_item_code(line):
            finish_current()
            current = None
            continue

        if current is None:
            continue

        if is_stop_line(line):
            finish_current()
            continue

        low = line.lower()
        if low in HEADER_WORDS:
            continue

        # A lone monetary value on the following line is a PDF extraction
        # artefact from Chadwicks totals/column alignment. It is NOT part
        # of the product description. Examples include 42.07, 29.66,
        # 117.90, 3534.80 and 1421.15.
        if re.fullmatch(r"\d{1,10}[.,]\d{2}", line):
            continue

        # Continuation lines are genuine description text. They are added
        # only after the numeric columns have already been removed from the
        # original item row.
        current["description"] = clean_text(
            f"{current['description']} {line}"
        )

    finish_current()

    # Safety fallback for an alternate layout where CODE/QTY/UNIT are on
    # separate lines.  This is deliberately conservative and is only used
    # when the real inline parser found nothing.
    if not rows:
        for i, line in enumerate(raw_lines):
            if not is_item_code(line):
                continue
            code = normalise_code(line)
            qty = None
            unit = None
            unit_idx = None
            for j in range(i + 1, min(i + 7, len(raw_lines))):
                q = re.fullmatch(r"(\d+(?:[.,]\d+)?)", raw_lines[j])
                if q and qty is None:
                    qty = to_price(q.group(1))
                    continue
                if qty is not None and is_unit(raw_lines[j]):
                    unit = raw_lines[j].upper()
                    unit_idx = j
                    break
            if qty is None or unit is None:
                continue

            desc = []
            prices = []
            k = unit_idx + 1
            while k < len(raw_lines):
                x = raw_lines[k]
                if is_item_code(x) or is_stop_line(x):
                    break
                ps = extract_prices(x)
                if ps:
                    prices.extend(ps)
                    # First price is unit price; stop after the table values.
                    if len(prices) >= 3:
                        break
                else:
                    desc.append(x)
                k += 1

            if prices and desc:
                rows.append({
                    "code": code,
                    "qty": qty,
                    "unit": unit,
                    "description": clean_text(" ".join(desc)),
                    "unit_price": prices[0],
                })

    # Remove duplicate rows while preserving order.
    unique = []
    seen = set()
    for row in rows:
        key = (
            row["code"], row["qty"], row["unit"],
            row["unit_price"], row["description"]
        )
        if key not in seen:
            seen.add(key)
            unique.append(row)

    return unique

def parse_invoice(pdf_path):
    text = find_pdf_text(pdf_path)

    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    invoice_number = extract_invoice_number(
        text,
        os.path.basename(pdf_path),
    )
    invoice_date = extract_invoice_date(text)

    items = parse_rows_from_lines(text)

    return {
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "items": items,
        "error": None,
    }


def find_sheet(wb, names):
    lower_map = {s.lower(): s for s in wb.sheetnames}
    for name in names:
        if name.lower() in lower_map:
            return wb[lower_map[name.lower()]]
    return None


def get_headers(ws):
    headers = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[clean_text(cell.value).lower()] = cell.column
    return headers


def find_column(headers, possible_names):
    for name in possible_names:
        if name.lower() in headers:
            return headers[name.lower()]
    return None


def load_master(workbook_path):
    wb = load_workbook(workbook_path)

    prices_ws = find_sheet(
        wb,
        ["MATERIAL_PRICES", "MATERIAL PRICES", "Material_Prices"],
    )

    master_ws = find_sheet(
        wb,
        ["MATERIAL_MASTER", "MATERIAL MASTER", "MATERIALS"],
    )

    history_ws = find_sheet(
        wb,
        ["PRICE_HISTORY", "PRICE HISTORY"],
    )

    if prices_ws is None:
        raise RuntimeError(
            "Could not find MATERIAL_PRICES sheet in workbook."
        )

    price_headers = get_headers(prices_ws)

    code_col = find_column(
        price_headers,
        ["Material ID", "Material Code", "Code", "Item Code"],
    )
    supplier_col = find_column(
        price_headers,
        ["Supplier", "Supplier Name"],
    )
    desc_col = find_column(
        price_headers,
        ["Material", "Description", "Material Name"],
    )
    unit_col = find_column(
        price_headers,
        ["Unit", "UOM"],
    )
    new_price_col = find_column(
        price_headers,
        ["New Price", "Current Price", "Price", "New Price Ex VAT (€)", "Latest Unit Price Ex VAT (€)"],
    )
    previous_col = find_column(
        price_headers,
        ["Previous Price", "Previous", "Previous Price Ex VAT (€)"],
    )
    last_invoice_col = find_column(
        price_headers,
        ["Last Invoice", "Invoice", "Last Invoice Number"],
    )
    last_updated_col = find_column(
        price_headers,
        ["Last Updated", "Updated", "Last Updated Date"],
    )

    if code_col is None or new_price_col is None:
        raise RuntimeError(
            "MATERIAL_PRICES must contain an Item Code/Material ID column and a New Price column (including New Price Ex VAT (€))."
        )

    if history_ws is None:
        history_ws = wb.create_sheet("PRICE_HISTORY")
        history_ws.append([
            "Date",
            "Supplier",
            "Material ID",
            "Material",
            "Old Price",
            "New Price",
            "Invoice",
        ])

    materials = {}

    for row in range(2, prices_ws.max_row + 1):
        code = normalise_code(prices_ws.cell(row, code_col).value)
        if not code:
            continue

        supplier = (
            clean_text(prices_ws.cell(row, supplier_col).value)
            if supplier_col
            else ""
        )

        key = (supplier.upper(), code)

        materials[key] = {
            "row": row,
            "code": code,
            "supplier": supplier,
            "description": (
                clean_text(prices_ws.cell(row, desc_col).value)
                if desc_col
                else ""
            ),
            "unit": (
                clean_text(prices_ws.cell(row, unit_col).value)
                if unit_col
                else ""
            ),
            "new_price_col": new_price_col,
            "previous_col": previous_col,
        }

    return wb, prices_ws, history_ws, materials, {
        "code_col": code_col,
        "supplier_col": supplier_col,
        "desc_col": desc_col,
        "unit_col": unit_col,
        "new_price_col": new_price_col,
        "previous_col": previous_col,
        "last_invoice_col": last_invoice_col,
        "last_updated_col": last_updated_col,
    }


def append_history(ws, invoice_date, supplier, code, description,
                   old_price, new_price, invoice_number):
    ws.append([
        invoice_date or datetime.now().strftime("%d/%m/%Y"),
        supplier,
        code,
        description,
        old_price,
        new_price,
        invoice_number,
    ])


def process_invoice(pdf_path, workbook_state):
    wb, prices_ws, history_ws, materials, cols = workbook_state

    parsed = parse_invoice(pdf_path)
    invoice_number = parsed["invoice_number"]
    invoice_date = parsed["invoice_date"]
    items = parsed["items"]

    if not items:
        return {
            "invoice": invoice_number,
            "file": os.path.basename(pdf_path),
            "items": 0,
            "updated": 0,
            "unchanged": 0,
            "new": 0,
            "review": 1,
            "error": parsed["error"] or "No complete invoice item rows detected",
            "details": [],
        }

    updated = 0
    unchanged = 0
    new_materials = 0
    review = 0
    details = []

    # Use one stable next-row counter for all restored materials in this run.
    # This guarantees that multiple deleted materials are restored in one run.
    next_restore_row = prices_ws.max_row + 1

    for item in items:
        code = item["code"]
        invoice_price = to_price(item["unit_price"])
        key = (SUPPLIER.upper(), code)

        material = materials.get(key)

        if material is None:
            # Material was deleted from the workbook, but exists on an invoice.
            # Restore it automatically using the invoice data.
            new_row = next_restore_row
            next_restore_row += 1
            if cols["code_col"]:
                prices_ws.cell(new_row, cols["code_col"]).value = code
            if cols["supplier_col"]:
                prices_ws.cell(new_row, cols["supplier_col"]).value = SUPPLIER
            if cols["desc_col"]:
                prices_ws.cell(new_row, cols["desc_col"]).value = item["description"]
            if cols["unit_col"]:
                prices_ws.cell(new_row, cols["unit_col"]).value = item["unit"]
            prices_ws.cell(new_row, cols["new_price_col"]).value = invoice_price
            if cols["previous_col"]:
                prices_ws.cell(new_row, cols["previous_col"]).value = None
            if cols["last_invoice_col"]:
                prices_ws.cell(new_row, cols["last_invoice_col"]).value = invoice_number
            if cols["last_updated_col"]:
                updated_cell = prices_ws.cell(new_row, cols["last_updated_col"])
                updated_cell.value = datetime.now()
                updated_cell.number_format = "dd/mm/yyyy hh:mm"

            # Add it to the in-memory index so a repeated occurrence in another
            # invoice during the same run is handled as an existing material.
            materials[key] = {
                "row": new_row,
                "code": code,
                "supplier": SUPPLIER,
                "description": item["description"],
                "unit": item["unit"],
                "new_price_col": cols["new_price_col"],
                "previous_col": cols["previous_col"],
            }

            append_history(
                history_ws, invoice_date, SUPPLIER, code,
                item["description"], None, invoice_price, invoice_number
            )

            new_materials += 1
            updated += 1
            details.append(
                f"RESTORED | {code} | {item['unit']} | "
                f"{item['description']} | €{invoice_price:.2f} | "
                f"ADDED BACK FROM INVOICE"
            )
            continue

        row = material["row"]
        old_price = to_price(
            prices_ws.cell(row, cols["new_price_col"]).value
        )

        # Update description/unit only if the workbook has these columns.
        if cols["desc_col"]:
            prices_ws.cell(
                row, cols["desc_col"]
            ).value = item["description"]

        if cols["unit_col"]:
            prices_ws.cell(
                row, cols["unit_col"]
            ).value = item["unit"]

        # The main row always records the invoice currently being processed
        # as the latest invoice for this material. This keeps Last Invoice
        # aligned with the current New Price when a later invoice changes it.
        if cols["last_invoice_col"]:
            prices_ws.cell(row, cols["last_invoice_col"]).value = invoice_number

        if cols["last_updated_col"]:
            updated_cell = prices_ws.cell(row, cols["last_updated_col"])
            updated_cell.value = datetime.now()
            updated_cell.number_format = "dd/mm/yyyy hh:mm"

        if old_price is None:
            prices_ws.cell(
                row, cols["new_price_col"]
            ).value = invoice_price

            if cols["previous_col"]:
                prices_ws.cell(
                    row, cols["previous_col"]
                ).value = None

            append_history(
                history_ws,
                invoice_date,
                SUPPLIER,
                code,
                item["description"],
                None,
                invoice_price,
                invoice_number,
            )

            updated += 1
            details.append(
                f"UPDATED | {code} | {item['unit']} | "
                f"{item['description']} | €{invoice_price:.2f}"
            )
            continue

        if abs(old_price - invoice_price) < 0.005:
            unchanged += 1
            details.append(
                f"UNCHANGED | {code} | {item['unit']} | "
                f"{item['description']} | €{invoice_price:.2f}"
            )
            continue

        # Price changed.
        if cols["previous_col"]:
            prices_ws.cell(
                row, cols["previous_col"]
            ).value = old_price

        prices_ws.cell(
            row, cols["new_price_col"]
        ).value = invoice_price

        append_history(
            history_ws,
            invoice_date,
            SUPPLIER,
            code,
            item["description"],
            old_price,
            invoice_price,
            invoice_number,
        )

        updated += 1

        details.append(
            f"UPDATED | {code} | {item['unit']} | "
            f"{item['description']} | €{old_price:.2f} -> €{invoice_price:.2f}"
        )

    return {
        "invoice": invoice_number,
        "file": os.path.basename(pdf_path),
        "items": len(items),
        "updated": updated,
        "unchanged": unchanged,
        "new": new_materials,
        "review": review,
        "error": None,
        "details": details,
    }


def main():
    if len(sys.argv) < 3:
        print(
            'Usage: python chadwicks_invoice_import_FINAL_FIXED_V2.py '
            '"CHAD_Invoices" "Chadwicks_Invoice_Material_Master.xlsx"'
        )
        sys.exit(1)

    invoice_folder = sys.argv[1]
    workbook_path = sys.argv[2]

    if not os.path.isdir(invoice_folder):
        print(f"ERROR: Invoice folder not found: {invoice_folder}")
        sys.exit(1)

    if not os.path.isfile(workbook_path):
        print(f"ERROR: Workbook not found: {workbook_path}")
        sys.exit(1)

    backup_path = os.path.splitext(workbook_path)[0] + "_BEFORE_FINAL_FIXED_V7_CLEAN.xlsx"
    report_path = os.path.splitext(workbook_path)[0] + "_FINAL_FIXED_V7_CLEAN_REPORT.csv"

    shutil.copy2(workbook_path, backup_path)

    workbook_state = load_master(workbook_path)

    pdfs = sorted(
        [
            os.path.join(invoice_folder, f)
            for f in os.listdir(invoice_folder)
            if f.lower().endswith(".pdf")
        ]
    )

    print(f"Found {len(pdfs)} PDF invoice(s).")

    total_items = 0
    total_updated = 0
    total_unchanged = 0
    total_new = 0
    total_review = 0
    errors = 0

    report_rows = []

    for idx, pdf_path in enumerate(pdfs, 1):
        print(f"\n[{idx}/{len(pdfs)}] {os.path.basename(pdf_path)}")

        try:
            result = process_invoice(pdf_path, workbook_state)

            if result["error"]:
                errors += 1
                print(f"  ERROR | {result['error']}")
            else:
                print(f"  Items detected: {result['items']}")

                for detail in result["details"]:
                    print(f"  {detail}")

            total_items += result["items"]
            total_updated += result["updated"]
            total_unchanged += result["unchanged"]
            total_new += result["new"]
            total_review += result["review"]

            report_rows.append({
                "File": result["file"],
                "Invoice": result["invoice"],
                "Items": result["items"],
                "Updated": result["updated"],
                "Unchanged": result["unchanged"],
                "New Materials": result["new"],
                "Needs Review": result["review"],
                "Error": result["error"] or "",
            })

        except Exception as exc:
            errors += 1
            print(f"  ERROR | {type(exc).__name__}: {exc}")

            report_rows.append({
                "File": os.path.basename(pdf_path),
                "Invoice": "",
                "Items": 0,
                "Updated": 0,
                "Unchanged": 0,
                "New Materials": 0,
                "Needs Review": 1,
                "Error": f"{type(exc).__name__}: {exc}",
            })

    wb = workbook_state[0]
    wb.save(workbook_path)

    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "File",
                "Invoice",
                "Items",
                "Updated",
                "Unchanged",
                "New Materials",
                "Needs Review",
                "Error",
            ],
        )
        writer.writeheader()
        writer.writerows(report_rows)

    print("\nFINAL FIXED V7 CLEAN TEST COMPLETE")
    print(f"PDFs found: {len(pdfs)}")
    print(f"Items detected: {total_items}")
    print(f"Prices updated: {total_updated}")
    print(f"Prices unchanged: {total_unchanged}")
    print(f"New materials: {total_new}")
    print(f"Needs review: {total_review}")
    print(f"Errors: {errors}")
    print(f"Backup: {os.path.basename(backup_path)}")
    print(f"Report: {os.path.basename(report_path)}")


if __name__ == "__main__":
    main()
