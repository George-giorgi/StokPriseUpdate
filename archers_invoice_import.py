import importlib.util
import os
import re


BASE_PATH = os.path.join(
    os.path.dirname(__file__),
    "chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py",
)

spec = importlib.util.spec_from_file_location("invoice_engine", BASE_PATH)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)

engine.SUPPLIER = "Archers"

ARCHERS_CODE_RE = re.compile(r"^\d{8}$")
ARCHERS_STOP_RE = re.compile(
    r"^(?:VAT Code|Net Goods|Net Delivery Charge|Your VAT ID|"
    r"Terms & Conditions|R e t u r n s)\b",
    re.I,
)
ARCHERS_NUMBER = (
    r"(?:\d{1,3}(?:,\d{3})*\.\d{2}|\d+(?:[.,]\d{2}))"
)


def _normalise_archers_text(text):
    text = text.replace("\r", "")
    # pdfplumber can split a decimal fraction onto its own line.
    text = re.sub(r"(?<=\d)\s*\n\s*(?=\.\d{2}\b)", "", text)
    # A wrapped description can run into the next item code on one line.
    text = re.sub(r"(?<!\d)\s+(?=\d{8}\s)", "\n", text)
    return text


def _clean_archers_description(description):
    description = engine.clean_text(description)
    description = re.sub(
        r"\s+(?:v|n|I|s|e|l|a|S)(?=\s|$)",
        " ",
        description,
    )
    return engine.clean_text(description)


def _parse_archers_rows(text):
    text = _normalise_archers_text(text)
    lines = [engine.clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line]

    rows = []
    current = []

    def finish_block(block):
        if not block:
            return

        block_text = engine.clean_text(" ".join(block))
        match = re.search(
            r"(?P<code>\d{8})\s+"
            r"(?P<description>.*?)\s+"
            rf"(?P<qty>{ARCHERS_NUMBER})\s+"
            r"(?:v|n)?(?P<unit>Each|EA|Box|Bag|Length|Pack|Roll|Set|Pair)\s+"
            rf"(?P<unit_price>{ARCHERS_NUMBER})\s+"
            rf"{ARCHERS_NUMBER}\s+"
            rf"{ARCHERS_NUMBER}\s+"
            rf"{ARCHERS_NUMBER}",
            block_text,
            re.I,
        )
        if not match:
            return

        description = _clean_archers_description(match.group("description"))
        trailing_text = _clean_archers_description(block_text[match.end():])
        if trailing_text and not ARCHERS_STOP_RE.match(trailing_text):
            description = _clean_archers_description(
                f"{description} {trailing_text}"
            )
        rows.append({
            "code": match.group("code"),
            "qty": engine.to_price(match.group("qty")),
            "unit": match.group("unit").upper(),
            "description": description,
            "unit_price": engine.to_price(match.group("unit_price")),
        })

    for line in lines:
        if ARCHERS_STOP_RE.match(line):
            finish_block(current)
            current = []
            continue

        if ARCHERS_CODE_RE.fullmatch(line.split()[0]) if line.split() else False:
            finish_block(current)
            current = [line]
        elif current:
            current.append(line)

    finish_block(current)

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


def parse_archers_invoice(pdf_path):
    text = engine.find_pdf_text(pdf_path)
    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    invoice_match = re.search(
        r"\bInvoice\s+No\.\s*([A-Z]?\d{6,})\b", text, re.I
    )
    date_match = re.search(
        r"\bInvoice Date\s+(\d{1,2}/\d{1,2}/\d{2,4}|\d{1,2}-\d{1,2}-\d{2,4})\b",
        text,
        re.I,
    )

    return {
        "invoice_number": (
            invoice_match.group(1).upper()
            if invoice_match
            else os.path.splitext(os.path.basename(pdf_path))[0]
        ),
        "invoice_date": date_match.group(1) if date_match else "",
        "items": _parse_archers_rows(text),
        "error": None,
    }


engine.parse_invoice = parse_archers_invoice


if __name__ == "__main__":
    engine.main()
