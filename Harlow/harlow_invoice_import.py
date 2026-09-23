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

engine.SUPPLIER = "Harlow"

ROW_RE = re.compile(
    r"^(?P<code>[A-Z0-9]+(?:[-/][A-Z0-9]+)*)\s+"
    r"(?P<description>.+?)\s+"
    r"(?P<qty>\d+(?:[.,]\d+)?)\s+"
    r"(?P<uom>[A-Z]+)\s+"
    r"(?P<unit_price>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)\s+"
    r"(?P<amount>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)\s+"
    r"(?P<discount>\d+(?:[.,]\d+)?)\s+"
    r"(?P<line_total>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)\s*$",
    re.I,
)


def _parse_harlow_rows(text):
    lines = [engine.clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line]
    rows = []
    in_items = False

    for line in lines:
        if line.startswith("Item Code Description"):
            in_items = True
            continue
        if not in_items:
            continue
        if line.startswith("Sales Invoice"):
            break

        match = ROW_RE.match(line)
        if not match:
            continue

        row = {
            "code": match.group("code").upper(),
            "qty": engine.to_price(match.group("qty")),
            "unit": "EA",
            "description": engine.clean_text(match.group("description")),
            "unit_price": engine.to_price(match.group("unit_price")),
            "amount": engine.to_price(match.group("line_total")),
        }
        rows.append(row)

    unique = []
    seen = set()
    for row in rows:
        key = (
            row["code"], row["qty"], row["unit_price"], row["description"], row["amount"]
        )
        if key not in seen:
            seen.add(key)
            unique.append(row)
    return unique


def parse_harlow_invoice(pdf_path):
    text = engine.find_pdf_text(pdf_path)
    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    header_match = re.search(
        r"\b[A-Z0-9]+\s+(?P<date>\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s+"
        r"(?P<invoice>[A-Z0-9]+)\s+(?:[A-Z0-9]+(?:/[A-Z0-9]+)?)\b",
        text,
        re.I,
    )
    invoice_number = (
        header_match.group("invoice").upper()
        if header_match
        else os.path.splitext(os.path.basename(pdf_path))[0]
    )
    invoice_date = (
        header_match.group("date")
        if header_match
        else ""
    )

    return {
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "items": _parse_harlow_rows(text),
        "error": None,
    }


def _ensure_harlow_workbook(workbook_path):
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
            f"Harlow workbook not found and template is missing: {workbook_path}"
        )

    os.makedirs(os.path.dirname(os.path.abspath(workbook_path)), exist_ok=True)
    shutil.copy2(template_path, workbook_path)
    print(f"Created Harlow workbook: {workbook_path}")


engine.parse_invoice = parse_harlow_invoice


if __name__ == "__main__":
    if len(os.sys.argv) >= 3:
        _ensure_harlow_workbook(os.sys.argv[2])
    engine.main()
