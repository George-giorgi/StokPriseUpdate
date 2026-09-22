import importlib.util
import os
import re
import shutil


BASE_PATH = os.path.join(
    os.path.dirname(__file__),
    "chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py",
)

spec = importlib.util.spec_from_file_location("invoice_engine", BASE_PATH)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)

engine.SUPPLIER = "Value"

VALUE_STOP_RE = re.compile(
    r"^(?:VAT Analysis|Subtotal|Bank Details|Note:|Total EUR|Total|Net Total|Delivery|DEL)\b",
    re.I,
)

VALUE_ITEM_RE = re.compile(
    r"^(?P<code>[A-Z0-9/_-]+)\s+(?P<rest>.+?)\s+"
    r"(?P<qty>\d{1,3}(?:[.,]\d{3})?|\d+(?:[.,]\d+)?)\s+"
    r"(?P<unit_price>\d{1,3}(?:,\d{3})*(?:\.\d{1,3})?|\d+(?:[.,]\d{1,3})?)\s+"
    r"(?P<amount>\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d+(?:[.,]\d{2})?)\s*$",
    re.I,
)


def _detect_value_unit(rest):
    lower = rest.lower()
    if "pallet" in lower:
        return "PALLET"
    if "units" in lower:
        return "UNITS"
    if "sheet" in lower:
        return "SHEET"
    if "roll" in lower:
        return "ROLL"
    return "EACH"


def _clean_value_description(rest):
    rest = engine.clean_text(rest)
    if not rest:
        return ""

    tokens = rest.split()
    if not tokens:
        return ""

    drop_words = {"DN", "DEL", "DELIVERY", "VOLCALIS", "NON", "RETURNABLE", "ITEM"}
    unit_words = {"PALLET", "UNITS", "ROLL", "SHEET", "UNIT", "UNITS"}

    while tokens and (tokens[-1].upper() in drop_words or tokens[-1].upper() in unit_words):
        tokens.pop()

    if not tokens:
        return ""

    if tokens[-1].upper() in {"N", "V", "I", "S", "E", "L", "A"}:
        tokens.pop()

    description = " ".join(tokens)
    description = re.sub(r"\s+[*]?[A-Z0-9]{8,}[*]?\s*$", "", description)
    description = re.sub(r"\s+(?:Volcalis|non returnable item|delivery)\s*$", "", description, flags=re.I)
    return engine.clean_text(description)


def _parse_value_block(block_text):
    if not block_text:
        return None

    match = VALUE_ITEM_RE.match(block_text)
    if not match:
        return None

    code = match.group("code").upper()
    rest = match.group("rest")
    qty = engine.to_price(match.group("qty"))
    unit_price = engine.to_price(match.group("unit_price"))
    amount = engine.to_price(match.group("amount"))

    description = _clean_value_description(rest)
    if not description or qty is None or unit_price is None:
        return None

    return {
        "code": code,
        "qty": qty,
        "unit": _detect_value_unit(rest),
        "description": description,
        "unit_price": unit_price,
        "amount": amount,
    }


def _parse_value_rows(text):
    lines = [engine.clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line]

    rows = []
    in_items = False

    for line in lines:
        if line.startswith("Item Code"):
            in_items = True
            continue

        if not in_items:
            continue

        if VALUE_STOP_RE.match(line):
            in_items = False
            continue

        if not line:
            continue

        if re.match(r"^[A-Z0-9/_-]+\s+", line) and not line.lower().startswith("vat"):
            item = _parse_value_block(line)
            if item is not None:
                rows.append(item)

    unique = []
    seen = set()
    for row in rows:
        key = (
            row["code"], row["qty"], row["unit"],
            row["unit_price"], row["description"],
        )
        if key not in seen:
            seen.add(key)
            unique.append(row)

    return unique


def parse_value_invoice(pdf_path):
    text = engine.find_pdf_text(pdf_path)
    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    invoice_match = re.search(
        r"^\s*(?P<invoice>\d{5,})\s+\d+\s+\S+\s+(?P<date>\d{1,2}/\d{1,2}/\d{2,4})\b",
        text,
        re.M,
    )
    invoice_number = invoice_match.group("invoice") if invoice_match else os.path.splitext(os.path.basename(pdf_path))[0]
    invoice_date = invoice_match.group("date") if invoice_match else ""

    items = _parse_value_rows(text)
    return {
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "items": items,
        "error": None,
    }


engine.parse_invoice = parse_value_invoice


def _ensure_value_workbook(workbook_path):
    if os.path.isfile(workbook_path):
        return

    template_path = os.path.join(
        os.path.dirname(__file__),
        "Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx",
    )
    if not os.path.isfile(template_path):
        raise FileNotFoundError(
            f"Value workbook not found and template is missing: {workbook_path}"
        )

    os.makedirs(os.path.dirname(os.path.abspath(workbook_path)), exist_ok=True)
    shutil.copy2(template_path, workbook_path)
    print(f"Created Value workbook: {workbook_path}")


if __name__ == "__main__":
    if len(os.sys.argv) >= 3:
        _ensure_value_workbook(os.sys.argv[2])
    engine.main()
