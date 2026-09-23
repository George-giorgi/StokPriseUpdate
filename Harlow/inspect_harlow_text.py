from pathlib import Path
import pdfplumber

root = Path(__file__).resolve().parent / 'Harlow_new_Invoices'
for pdf_path in sorted(root.glob('*.pdf')):
    print(f'\n==== {pdf_path.name} ====')
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages[:4], 1):
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ''
            print(f'--- PAGE {i} ---')
            print(text[:8000])
            print()
