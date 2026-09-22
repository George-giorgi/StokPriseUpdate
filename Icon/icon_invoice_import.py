import importlib.util
import os
import re
import shutil


BASE_PATH = os.path.join(
    os.path.dirname(__file__),
    "..",
    "Chadwicks",
    "chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py",
)

spec = importlib.util.spec_from_file_location("invoice_engine", BASE_PATH)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)

engine.SUPPLIER = "Icon"

STOP_RE = re.compile(
    r"^(?:Delivery Address|Total Discount|Total Net Amount|Carriage Net|"
    r"Total VAT Amount|Invoice Total|Wire Transfer Instructions)\b",
    re.I,
)

ROW_RE = re.compile(
    r"^(?P<qty>\d+(?:[.,]\d{3})?)\s+"
    r"(?P<code>[A-Z0-9]+(?:[-/][A-Z0-9]+)*)\s+"
    r"(?P<description>.+?)\s+"
    r"(?P<unit_price>\d+(?:[.,]\d{3})?)\s+"
    r"(?P<discount>\d+(?:[.,]\d{2})?)\s+"
    r"(?P<net>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)\s+"
    r"(?P<vat_rate>\d+(?:[.,]\d{2})?)\s+"
    r"(?P<vat_amount>\d+(?:[.,]\d{2})?)$",
    re.I,
)


def _parse_icon_rows(text):
    lines = [engine.clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line]
    rows = []
    in_items = False
    current = None

    def finish_current():
        nonlocal current
        if current:
            rows.append(current)
        current = None

    for line in lines:
        if line.startswith("Quantity Details"):
            in_items = True
            continue
        if not in_items:
            continue
        if STOP_RE.match(line):
            finish_current()
            break

        match = ROW_RE.match(line)
        if match:
            finish_current()
            current = {
                "code": match.group("code").upper(),
                "qty": engine.to_price(match.group("qty")),
                "unit": "EA",
                "description": engine.clean_text(match.group("description")),
                "unit_price": engine.to_price(match.group("unit_price")),
                "amount": engine.to_price(match.group("net")),
            }
        elif current:
            current["description"] = engine.clean_text(
                f"{current['description']} {line}"
            )

    finish_current()

    unique = []
    seen = set()
    for row in rows:
        key = (
            row["code"], row["qty"], row["unit_price"],
            row["amount"], row["description"],
        )
        if key not in seen:
            seen.add(key)
            unique.append(row)
    return unique


def parse_icon_invoice(pdf_path):
    text = engine.find_pdf_text(pdf_path)
    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    invoice_match = re.search(r"\bInvoice\s+(\d+)\b", text, re.I)
    date_match = re.search(
        r"\bDate\s+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", text, re.I
    )
    return {
        "invoice_number": (
            invoice_match.group(1)
            if invoice_match
            else os.path.splitext(os.path.basename(pdf_path))[0]
        ),
        "invoice_date": date_match.group(1) if date_match else "",
        "items": _parse_icon_rows(text),
        "error": None,
    }


def _ensure_icon_workbook(workbook_path):
    if os.path.isfile(workbook_path):
        return

    template_path = os.path.join(
        os.path.dirname(__file__),
        "..",
        "Chadwicks",
        "Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx",
    )
    if not os.path.isfile(template_path):
        raise FileNotFoundError(
            f"Icon workbook not found and template is missing: {workbook_path}"
        )

    os.makedirs(os.path.dirname(os.path.abspath(workbook_path)), exist_ok=True)
    shutil.copy2(template_path, workbook_path)
    print(f"Created Icon workbook: {workbook_path}")


engine.parse_invoice = parse_icon_invoice


if __name__ == "__main__":
    if len(os.sys.argv) >= 3:
        _ensure_icon_workbook(os.sys.argv[2])
    engine.main()
