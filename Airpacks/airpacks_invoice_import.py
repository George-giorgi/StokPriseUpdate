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

engine.SUPPLIER = "Airpacks"

STOP_RE = re.compile(
    r"^(?:Net Amount|VAT Anaysis|VAT Analysis|Total Gross|All Amounts|VAT Number)\b",
    re.I,
)

ROW_TAIL_RE = re.compile(
    r"(?P<qty>\d+(?:[.,]\d+)?)\s+"
    r"(?P<unit>Linn\s+Meter|Each|EA|Sheet|Box|Pack|Roll|Set|Bag|Length|Unit)\s+"
    r"(?P<unit_price>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)\s+"
    r"(?P<amount>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)",
    re.I,
)

CODE_PREFIXES = (
    re.compile(r"^(?P<code>Pipe Lagging Size\s+\d+)\s+(?P<description>.+)$", re.I),
    re.compile(r"^(?P<code>D Loft Hooks and Eye)\s+(?P<description>.+)$", re.I),
    re.compile(r"^(?P<code>Hot Water Tank Size\d+)\s+(?P<description>.+)$", re.I),
    re.compile(r"^(?P<code>[A-Z0-9][A-Z0-9./&-]*/[A-Z0-9./&-]+)\s+(?P<description>.+)$", re.I),
    re.compile(r"^(?P<code>C[A-Z0-9/&-]+)\s+(?P<description>.+)$", re.I),
    re.compile(r"^(?P<code>XPS\S+|EWI\S+|D&W|[A-Z])\s+(?P<description>.+)$", re.I),
)


def _normalise_airpacks_text(text):
    text = text.replace("\r", "")
    text = re.sub(r"(?<=\d)\s*\n\s*(?=\.\d{2}\b)", "", text)
    return text


def _parse_airpacks_row(row_text):
    row_text = engine.clean_text(row_text)
    match = ROW_TAIL_RE.search(row_text)
    if not match:
        return None

    prefix = engine.clean_text(row_text[:match.start()])
    code = ""
    description = prefix
    for code_pattern in CODE_PREFIXES:
        code_match = code_pattern.match(prefix)
        if code_match:
            code = engine.clean_text(code_match.group("code")).upper()
            description = engine.clean_text(code_match.group("description"))
            break

    if not code or not description:
        return None

    if code == "C":
        tank_lid_match = re.match(r"(Water Tank Lid\s+\d+)", description, re.I)
        if tank_lid_match:
            code = f"C {tank_lid_match.group(1)}"

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


def _parse_airpacks_rows(text):
    text = _normalise_airpacks_text(text)
    lines = [engine.clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line]

    rows = []
    in_items = False
    current = []

    def finish_current():
        nonlocal current
        if not current:
            return
        item = _parse_airpacks_row(" ".join(current))
        if item:
            rows.append(item)
        current = []

    for line in lines:
        if line.startswith("Item Code"):
            in_items = True
            continue
        if not in_items:
            continue
        if STOP_RE.match(line):
            finish_current()
            break

        if current and any(pattern.match(line) for pattern in CODE_PREFIXES):
            finish_current()

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


def _parse_molloy_rows(text):
    rows = []
    money_re = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")
    unit_re = re.compile(r"Each|EA|M3|Ton|Load", re.I)

    for line in text.splitlines():
        line = engine.clean_text(line)
        if not line or not re.match(r"^[A-Z0-9]+\s+", line):
            continue

        match = re.match(
            r"^(?P<code>[A-Z0-9]+)\s+(?P<description>.*?)\s+"
            r"(?P<qty>\d+(?:[.,]\d+)?)\s+(?P<unit>Each|EA|M3|Ton|Load)\s+"
            r"(?P<tail>.+)$",
            line,
            re.I,
        )
        if not match:
            continue

        prices = money_re.findall(match.group("tail"))
        if len(prices) < 2:
            continue

        rows.append({
            "code": match.group("code").upper(),
            "qty": engine.to_price(match.group("qty")),
            "unit": match.group("unit").upper(),
            "description": engine.clean_text(match.group("description")),
            "unit_price": engine.to_price(prices[0]),
            "amount": engine.to_price(prices[-1]),
        })

    return rows


def parse_airpacks_invoice(pdf_path):
    text = engine.find_pdf_text(pdf_path)
    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    invoice_match = re.search(
        r"\b(?:Invoice|Sales Invoice Number)\s*(?:No\.?)?\s*(SI-\d+)\b",
        text,
        re.I,
    )
    date_match = re.search(
        r"\bDate:\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", text, re.I
    )
    if "Molloy Concrete Ltd" in text:
        items = []
        parse_error = "This is a Molloy Concrete invoice; move it to a Molloy invoice folder."
    else:
        items = _parse_airpacks_rows(text)
        parse_error = None

    return {
        "invoice_number": (
            invoice_match.group(1).upper()
            if invoice_match
            else os.path.splitext(os.path.basename(pdf_path))[0]
        ),
        "invoice_date": date_match.group(1) if date_match else "",
        "items": items,
        "error": parse_error,
    }


def _ensure_airpacks_workbook(workbook_path):
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
            f"Airpacks workbook not found and template is missing: {workbook_path}"
        )

    os.makedirs(os.path.dirname(os.path.abspath(workbook_path)), exist_ok=True)
    shutil.copy2(template_path, workbook_path)
    print(f"Created Airpacks workbook: {workbook_path}")


engine.parse_invoice = parse_airpacks_invoice


if __name__ == "__main__":
    if len(os.sys.argv) >= 3:
        _ensure_airpacks_workbook(os.sys.argv[2])
    engine.main()
