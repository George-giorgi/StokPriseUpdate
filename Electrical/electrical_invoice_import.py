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

engine.SUPPLIER = "Electrical"

STOP_RE = re.compile(
    r"^(?:CONDITIONS OF SALE|SUB-TOTAL|WEEE:|Total charges|VAT\b|Total\b|Where applicable)\b",
    re.I,
)

ROW_TAIL_RE = re.compile(
    r"(?P<qty>\d+(?:[.,]\d+)?)\s+"
    r"(?P<unit_price>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)\s+"
    r"(?P<unit>EA|Each|Unit|Pack|Box|M|MTR)\s+"
    r"(?P<vat>\d+(?:[.,]\d+)?\s*%)\s+"
    r"(?P<weee>\d+(?:[.,]\d+)?)\s+"
    r"(?P<amount>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)$",
    re.I,
)


CODE_PREFIX_RE = re.compile(
    r"^(?P<code>\d{5,})\s+(?P<config>\S+)\s+(?P<description>.+)$",
    re.I,
)


def _normalise_electrical_text(text):
    return text.replace("\r", "")


def _parse_electrical_row(row_text):
    row_text = engine.clean_text(row_text)
    match = ROW_TAIL_RE.search(row_text)
    if not match:
        return None

    prefix = engine.clean_text(row_text[:match.start()])
    code_match = CODE_PREFIX_RE.match(prefix)
    if not code_match:
        return None

    code = code_match.group("code")
    config = code_match.group("config")
    description = engine.clean_text(
        f"{config} {code_match.group('description')}"
    )
    qty = engine.to_price(match.group("qty"))
    unit_price = engine.to_price(match.group("unit_price"))
    amount = engine.to_price(match.group("amount"))
    if qty is None or unit_price is None:
        return None

    return {
        "code": code,
        "qty": qty,
        "unit": match.group("unit").upper(),
        "description": description,
        "unit_price": unit_price,
        "amount": amount,
    }


def _parse_electrical_rows(text):
    lines = [engine.clean_text(line) for line in _normalise_electrical_text(text).splitlines()]
    lines = [line for line in lines if line]

    rows = []
    in_items = False
    current = []

    def finish_current():
        nonlocal current
        if not current:
            return
        item = _parse_electrical_row(" ".join(current))
        if item:
            rows.append(item)
        current = []

    for line in lines:
        if line.startswith("Item Number"):
            in_items = True
            continue
        if not in_items:
            continue
        if STOP_RE.match(line):
            finish_current()
            break

        if current and re.match(r"^\d{5,}\s+", line):
            finish_current()

        if not current and rows and not re.match(r"^\d{5,}\s+", line):
            rows[-1]["description"] = engine.clean_text(
                f"{rows[-1]['description']} {line}"
            )
            continue

        current.append(line)
        if ROW_TAIL_RE.search(" ".join(current)):
            finish_current()

    finish_current()

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


def parse_electrical_invoice(pdf_path):
    text = engine.find_pdf_text(pdf_path)
    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    invoice_match = re.search(
        r"\bInvoice Number\s+(RSI\d+)\b", text, re.I
    )
    date_match = re.search(
        r"\bDate\s+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", text, re.I
    )
    return {
        "invoice_number": (
            invoice_match.group(1).upper()
            if invoice_match
            else os.path.splitext(os.path.basename(pdf_path))[0]
        ),
        "invoice_date": date_match.group(1) if date_match else "",
        "items": _parse_electrical_rows(text),
        "error": None,
    }


def _ensure_electrical_workbook(workbook_path):
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
            f"Electrical workbook not found and template is missing: {workbook_path}"
        )

    os.makedirs(os.path.dirname(os.path.abspath(workbook_path)), exist_ok=True)
    shutil.copy2(template_path, workbook_path)
    print(f"Created Electrical workbook: {workbook_path}")


engine.parse_invoice = parse_electrical_invoice


if __name__ == "__main__":
    if len(os.sys.argv) >= 3:
        _ensure_electrical_workbook(os.sys.argv[2])
    engine.main()
