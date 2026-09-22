import importlib.util
import os
import re
import shutil

import pdfplumber


BASE_PATH = os.path.join(
    os.path.dirname(__file__),
    "..",
    "Chadwicks",
    "chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py",
)

spec = importlib.util.spec_from_file_location("invoice_engine", BASE_PATH)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)

engine.SUPPLIER = "KMS"

STOP_RE = re.compile(
    r"^(?:VAT Rate|TOTAL|Terms and Conditions|KMS CONCEPTS|Registered Address)\b",
    re.I,
)

ROW_TAIL_RE = re.compile(
    r"(?P<qty>\d+(?:[.,]\d+)?)\s+"
    r"(?P<unit_price>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)\s+"
    r"(?P<vat>\d+(?:[.,]\d+)?\s*%?)\s+"
    r"(?P<amount>\d{1,3}(?:,\d{3})*(?:[.,]\d{2})?)$",
    re.I,
)

ROW_START_RE = re.compile(
    r"^(?:[A-Z0-9]+(?:[-/][A-Z0-9.]+)*|[A-Z]+\s+[A-Z]+)\s+",
    re.I,
)


def _parse_kms_row(row_text):
    row_text = engine.clean_text(row_text)
    match = ROW_TAIL_RE.search(row_text)
    if not match:
        return None

    prefix = engine.clean_text(row_text[:match.start()])
    code_match = re.match(r"^(\S+)(?:\s+|$)(.*)$", prefix)
    if not code_match:
        return None

    code = code_match.group(1).upper()
    description = engine.clean_text(code_match.group(2))
    if not description:
        return None

    return {
        "code": code,
        "qty": engine.to_price(match.group("qty")),
        "unit": "EA",
        "description": description,
        "unit_price": engine.to_price(match.group("unit_price")),
        "amount": engine.to_price(match.group("amount")),
    }


def _extract_kms_rows(pdf_path):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=2, y_tolerance=3)
            grouped = {}
            for word in words:
                grouped.setdefault(round(float(word["top"]), 1), []).append(word)

            in_items = False
            current = None
            description_x = 120
            quantity_x = 340
            has_code_column = False

            def finish_current():
                nonlocal current
                if current and current["code"].upper() not in {"CARRIAGE", "PALLET"}:
                    rows.append(current)
                current = None

            for top in sorted(grouped):
                line_words = sorted(grouped[top], key=lambda word: word["x0"])
                line = engine.clean_text(" ".join(word["text"] for word in line_words))
                if line.startswith("Code") and "Description" in line:
                    in_items = True
                    has_code_column = True
                    description_words = [
                        word for word in line_words if word["text"].lower() == "description"
                    ]
                    quantity_words = [
                        word for word in line_words if word["text"].lower() in {"qty/hrs", "qty"}
                    ]
                    if description_words:
                        description_x = float(description_words[0]["x0"])
                    if quantity_words:
                        quantity_x = float(quantity_words[0]["x0"]) - 15
                    continue
                if line.startswith("Description"):
                    in_items = True
                    has_code_column = False
                    description_x = float(line_words[0]["x0"])
                    continue
                if not in_items:
                    continue
                if STOP_RE.match(line):
                    finish_current()
                    break

                if line.startswith("Pallet Delivery") or line.startswith("Carriage"):
                    finish_current()
                    continue

                left = [word["text"] for word in line_words if float(word["x0"]) < description_x]
                middle = [word["text"] for word in line_words if description_x <= float(word["x0"]) < quantity_x]
                right = [word["text"] for word in line_words if float(word["x0"]) >= quantity_x]
                numeric = ROW_TAIL_RE.search(" ".join(right))

                if numeric and left:
                    finish_current()
                    if has_code_column:
                        code = engine.clean_text(" ".join(left))
                        description = engine.clean_text(" ".join(middle))
                    else:
                        code = engine.clean_text(" ".join(left + middle))
                        description = code
                    current = {
                        "code": re.sub(r"(?<=-)\s+", "", code),
                        "qty": engine.to_price(numeric.group("qty")),
                        "unit": "EA",
                        "description": description,
                        "unit_price": engine.to_price(numeric.group("unit_price")),
                        "amount": engine.to_price(numeric.group("amount")),
                    }
                elif numeric and not left:
                    finish_current()
                    description = engine.clean_text(" ".join(middle))
                    current = {
                        "code": description,
                        "qty": engine.to_price(numeric.group("qty")),
                        "unit": "EA",
                        "description": description,
                        "unit_price": engine.to_price(numeric.group("unit_price")),
                        "amount": engine.to_price(numeric.group("amount")),
                    }
                elif current:
                    if left:
                        current["code"] = re.sub(
                            r"(?<=-)\s+", "", engine.clean_text(
                                f"{current['code']} {' '.join(left)}"
                            )
                        )
                    if middle:
                        if middle[0].upper() in {"COMMODITY", "COUNTRY"}:
                            continue
                        current["description"] = engine.clean_text(
                            f"{current['description']} {' '.join(middle)}"
                        )

            finish_current()

    unique = []
    seen = set()
    for row in rows:
        description = row["description"].upper()
        if "EWI EPS DISC GREY 68MM" in description:
            row["code"] = "EWI EPS Disc Grey 68mm"
        elif "HAMMER PLUG 6 X 60MM" in description:
            row["code"] = "236-5907704407249"
        elif "2025 ZAKU EWI SAF 30 WINDOW PROTECTION CLEAR FOIL" in description:
            row["code"] = "2025 ZAKU EWI SAF 30 WINDOW PROTECTION CLEAR FOIL"
        elif "134-SHIMS 15MM" in row["code"].upper():
            size = re.search(r"\b(3\s*MM|5\s*MM)\b", description)
            if size:
                row["code"] = f"134-shims 15mm {size.group(1).replace(' ', '')}"

        key = (
            row["code"], row["qty"], row["unit_price"], row["description"]
        )
        if key not in seen:
            seen.add(key)
            unique.append(row)
    return unique


def parse_kms_invoice(pdf_path):
    text = engine.find_pdf_text(pdf_path)
    if not text.strip():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "PDF text could not be extracted",
        }

    if "SALES CREDIT NOTE" in text.upper():
        return {
            "invoice_number": os.path.splitext(os.path.basename(pdf_path))[0],
            "invoice_date": "",
            "items": [],
            "error": "Credit note skipped; no price update applied",
        }

    invoice_match = re.search(r"\bInvoice Number.*?\b(SE-\d+)\b", text, re.I | re.S)
    date_match = re.search(
        r"\bInvoice Date.*?(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b",
        text,
        re.I | re.S,
    )
    return {
        "invoice_number": (
            invoice_match.group(1).upper()
            if invoice_match
            else os.path.splitext(os.path.basename(pdf_path))[0]
        ),
        "invoice_date": date_match.group(1) if date_match else "",
        "items": _extract_kms_rows(pdf_path),
        "error": None,
    }


def _ensure_kms_workbook(workbook_path):
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
            f"KMS workbook not found and template is missing: {workbook_path}"
        )

    os.makedirs(os.path.dirname(os.path.abspath(workbook_path)), exist_ok=True)
    shutil.copy2(template_path, workbook_path)
    print(f"Created KMS workbook: {workbook_path}")


engine.parse_invoice = parse_kms_invoice

_engine_process_invoice = engine.process_invoice


def _process_kms_invoice(pdf_path, workbook_state):
    parsed = parse_kms_invoice(pdf_path)
    if parsed["error"] and "Credit note" in parsed["error"]:
        return {
            "invoice": parsed["invoice_number"],
            "file": os.path.basename(pdf_path),
            "items": 0,
            "updated": 0,
            "unchanged": 0,
            "new": 0,
            "review": 0,
            "error": None,
            "details": ["SKIPPED | Credit note; no price update applied"],
            "over_10_changes": [],
        }
    return _engine_process_invoice(pdf_path, workbook_state)


engine.process_invoice = _process_kms_invoice


if __name__ == "__main__":
    if len(os.sys.argv) >= 3:
        _ensure_kms_workbook(os.sys.argv[2])
    engine.main()
