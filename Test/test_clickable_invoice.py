from pathlib import Path
import csv
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import importlib.util

ROOT = Path(__file__).resolve().parent.parent
CHADWICKS_ENGINE = ROOT / 'Chadwicks' / 'chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py'
TEMPLATE_WORKBOOK = ROOT / 'Chadwicks' / 'Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx'
DEFAULT_TEST_FOLDER = ROOT / 'Test' / 'Test_new_Invoices'
DEFAULT_WORKBOOK = ROOT / 'Test' / 'Test_MATERIAL_PRICES.xlsx'


def load_engine():
    spec = importlib.util.spec_from_file_location('test_engine', CHADWICKS_ENGINE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.SUPPLIER = 'Test'
    return module


def ensure_workbook(workbook_path: Path):
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    if not workbook_path.exists() and TEMPLATE_WORKBOOK.exists():
        shutil.copy2(TEMPLATE_WORKBOOK, workbook_path)
    return workbook_path


def run_processing(invoice_folder: str, workbook_path: str):
    folder = Path(invoice_folder)
    workbook = Path(workbook_path)

    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f'Invoice folder not found: {folder}')

    ensure_workbook(workbook)
    if not workbook.exists():
        raise FileNotFoundError(f'Workbook not found or could not be created: {workbook}')

    engine = load_engine()
    workbook_state = engine.load_master(str(workbook))
    pdfs = sorted(
        [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == '.pdf']
    )

    if not pdfs:
        return {
            'pdf_count': 0,
            'items': 0,
            'updated': 0,
            'unchanged': 0,
            'new': 0,
            'report_path': '',
            'email_status': 'No PDFs found',
        }

    # Keep the same workbook file; update prices in-place.
    backup_path = workbook.with_name(workbook.stem + '_BEFORE_FINAL_FIXED_V7_CLEAN.xlsx')
    if not backup_path.exists():
        shutil.copy2(workbook, backup_path)

    total_items = 0
    total_updated = 0
    total_unchanged = 0
    total_new = 0
    total_review = 0
    errors = 0
    report_rows = []
    over_10_changes = []

    for pdf_path in pdfs:
        try:
            result = engine.process_invoice(str(pdf_path), workbook_state)
        except Exception as exc:
            errors += 1
            report_rows.append({
                'File': pdf_path.name,
                'Invoice': '',
                'Items': 0,
                'Updated': 0,
                'Unchanged': 0,
                'New Materials': 0,
                'Needs Review': 1,
                'Error': f'{type(exc).__name__}: {exc}',
            })
            continue

        report_rows.append({
            'File': result['file'],
            'Invoice': result['invoice'],
            'Items': result['items'],
            'Updated': result['updated'],
            'Unchanged': result['unchanged'],
            'New Materials': result['new'],
            'Needs Review': result['review'],
            'Error': result['error'] or '',
        })

        if result.get('error'):
            errors += 1
        else:
            total_items += result.get('items', 0)
            total_updated += result.get('updated', 0)
            total_unchanged += result.get('unchanged', 0)
            total_new += result.get('new', 0)
            total_review += result.get('review', 0)
            over_10_changes.extend(result.get('over_10_changes', []))

    workbook_state[0].save(str(workbook))

    report_path = workbook.with_name(workbook.stem + '_FINAL_FIXED_V7_CLEAN_REPORT.csv')
    with open(report_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                'File',
                'Invoice',
                'Items',
                'Updated',
                'Unchanged',
                'New Materials',
                'Needs Review',
                'Error',
            ],
        )
        writer.writeheader()
        writer.writerows(report_rows)

    try:
        engine.send_email_report(
            len(pdfs),
            total_items,
            total_updated,
            total_unchanged,
            total_new,
            total_review,
            errors,
            str(report_path),
            over_10_changes,
        )
        email_status = 'SENT'
    except Exception as exc:
        email_status = f'NOT SENT | {type(exc).__name__}: {exc}'

    return {
        'pdf_count': len(pdfs),
        'items': total_items,
        'updated': total_updated,
        'unchanged': total_unchanged,
        'new': total_new,
        'report_path': str(report_path),
        'email_status': email_status,
    }


def pick_folder_and_run():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    folder = filedialog.askdirectory(
        initialdir=str(DEFAULT_TEST_FOLDER),
        title='Select the invoice folder',
    )
    if not folder:
        return

    workbook = str(DEFAULT_WORKBOOK)
    if not DEFAULT_WORKBOOK.exists():
        workbook = filedialog.asksaveasfilename(
            initialfile='Test_MATERIAL_PRICES.xlsx',
            defaultextension='.xlsx',
            initialdir=str(ROOT / 'Test'),
            title='Choose workbook name and location',
        )
        if not workbook:
            workbook = str(DEFAULT_WORKBOOK)

    try:
        result = run_processing(folder, workbook)
        messagebox.showinfo(
            'Processing complete',
            'Invoices processed successfully.\n\n'
            f'PDFs: {result["pdf_count"]}\n'
            f'Items found: {result["items"]}\n'
            f'Updated: {result["updated"]}\n'
            f'Unchanged: {result["unchanged"]}\n'
            f'New/restored: {result["new"]}\n'
            f'Email: {result["email_status"]}\n\n'
            f'Workbook: {workbook}\n'
            f'Report: {result["report_path"]}'
        )
    except Exception as exc:  # pragma: no cover - GUI error path
        messagebox.showerror('Processing failed', str(exc))

    root.destroy()


if __name__ == '__main__':
    pick_folder_and_run()
