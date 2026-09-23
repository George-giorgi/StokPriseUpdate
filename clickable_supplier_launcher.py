from pathlib import Path
import csv
import importlib.util
import shutil
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

ROOT = Path(__file__).resolve().parent

SUPPLIERS = {
    'Airpacks': {
        'folder': 'Airpacks',
        'new_invoices': 'Airpacks_new_Invoices',
        'workbook': 'Airpacks_MATERIAL_PRICES.xlsx',
        'parser': 'Airpacks/airpacks_invoice_import.py',
    },
    'Archers': {
        'folder': 'Archers',
        'new_invoices': 'Archers_new_Invoices',
        'workbook': 'Archers_MATERIAL_PRICES.xlsx',
        'parser': 'Archers/archers_invoice_import.py',
    },
    'Chadwicks': {
        'folder': 'Chadwicks',
        'new_invoices': 'Chadwicks_new_Invoices',
        'workbook': 'Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx',
        'parser': 'Chadwicks/chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py',
    },
    'Electrical': {
        'folder': 'Electrical',
        'new_invoices': 'Electrical_new_Invoices',
        'workbook': 'Electrical_MATERIAL_PRICES.xlsx',
        'parser': 'Electrical/electrical_invoice_import.py',
    },
    'Harlow': {
        'folder': 'Harlow',
        'new_invoices': 'Harlow_new_Invoices',
        'workbook': 'Harlow_MATERIAL_PRICES.xlsx',
        'parser': 'Harlow/harlow_invoice_import.py',
    },
    'Icon': {
        'folder': 'Icon',
        'new_invoices': 'Icon_new_Invoices',
        'workbook': 'Icon_MATERIAL_PRICES.xlsx',
        'parser': 'Icon/icon_invoice_import.py',
    },
    'KMS': {
        'folder': 'KMS',
        'new_invoices': 'KMS_new_Invoices',
        'workbook': 'KMS_MATERIAL_PRICES.xlsx',
        'parser': 'KMS/kms_invoice_import.py',
    },
    'Southeren': {
        'folder': 'Southeren',
        'new_invoices': 'Southeren_new_Invoices',
        'workbook': 'Southern_MATERIAL_PRICES.xlsx',
        'parser': 'Southeren/southern_invoice_import.py',
    },
    'Test': {
        'folder': 'Test',
        'new_invoices': 'Test_new_Invoices',
        'workbook': 'Test_MATERIAL_PRICES.xlsx',
        'parser': 'Test/test_clickable_invoice.py',
    },
    'Value': {
        'folder': 'Value',
        'new_invoices': 'Value_new_Invoices',
        'workbook': 'Value_MATERIAL_PRICES.xlsx',
        'parser': 'Value/value_invoice_import.py',
    },
}


def supplier_names():
    return sorted(SUPPLIERS)


def supplier_root(supplier_name: str) -> Path:
    config = SUPPLIERS[supplier_name]
    return ROOT / config['folder']


def default_invoice_folder(supplier_name: str) -> Path:
    config = SUPPLIERS[supplier_name]
    return supplier_root(supplier_name) / config['new_invoices']


def default_workbook(supplier_name: str) -> Path:
    config = SUPPLIERS[supplier_name]
    return supplier_root(supplier_name) / config['workbook']


def load_supplier_module(supplier_name: str):
    config = SUPPLIERS[supplier_name]
    parser_path = ROOT / config['parser']
    if not parser_path.exists():
        raise FileNotFoundError(f'Parser not found for {supplier_name}: {parser_path}')

    spec = importlib.util.spec_from_file_location(f'{supplier_name}_engine', parser_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if hasattr(module, 'engine'):
        module.engine.SUPPLIER = supplier_name
    return module


def ensure_workbook(workbook_path: Path, supplier_name: str):
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    template_path = ROOT / 'Chadwicks' / 'Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx'
    if not workbook_path.exists() and template_path.exists():
        shutil.copy2(template_path, workbook_path)
    return workbook_path


def validate_supplier_folder(folder: Path, supplier_name: str):
    expected = supplier_root(supplier_name).resolve()
    resolved = folder.resolve()
    if resolved == expected or expected in resolved.parents:
        return

    raise ValueError(
        f'Wrong supplier folder selected.\n\nThis launcher is for {supplier_name}.\n'
        f'Please choose a folder inside: {expected}\nYou selected: {folder}'
    )


def run_processing(supplier_name: str, invoice_folder: str, workbook_path: str):
    supplier = supplier_name
    folder = Path(invoice_folder)
    workbook = Path(workbook_path)

    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f'Invoice folder not found: {folder}')

    validate_supplier_folder(folder, supplier)

    pdfs = sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == '.pdf'])
    if not pdfs:
        raise ValueError(
            f'No PDF invoice files were found in this folder:\n{folder}\n\n'
            'Please choose the correct invoice folder.'
        )

    ensure_workbook(workbook, supplier)
    if not workbook.exists():
        raise FileNotFoundError(f'Workbook not found or could not be created: {workbook}')

    module = load_supplier_module(supplier)
    engine = module.engine if hasattr(module, 'engine') else module
    workbook_state = engine.load_master(str(workbook))

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


def pick_folder_and_run(supplier_name: str):
    supplier = supplier_name
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    folder = filedialog.askdirectory(
        initialdir=str(default_invoice_folder(supplier)),
        title=f'Select {supplier} invoice folder',
    )
    if not folder:
        return

    try:
        validate_supplier_folder(Path(folder), supplier)
    except ValueError as exc:
        messagebox.showerror('Wrong supplier folder', str(exc))
        return

    workbook = str(default_workbook(supplier))
    if not default_workbook(supplier).exists():
        workbook = filedialog.asksaveasfilename(
            initialfile=default_workbook(supplier).name,
            defaultextension='.xlsx',
            initialdir=str(supplier_root(supplier)),
            title=f'Choose workbook for {supplier}',
        )
        if not workbook:
            workbook = str(default_workbook(supplier))

    loading = tk.Toplevel(root)
    loading.title(f'Processing {supplier} invoices')
    loading.geometry('340x110')
    loading.resizable(False, False)
    loading.attributes('-topmost', True)

    label = ttk.Label(loading, text='Processing PDFs and updating prices...')
    label.pack(pady=(18, 8))

    spinner = ttk.Progressbar(loading, mode='indeterminate')
    spinner.pack(fill='x', padx=20)
    spinner.start(18)
    loading.update_idletasks()

    result = {}
    error = {}

    def worker():
        try:
            result.update(run_processing(supplier, folder, workbook))
        except Exception as exc:
            error['value'] = exc
        finally:
            loading.after(0, loading.destroy)
            root.after(0, lambda: _finish_processing(result, error))

    def _finish_processing(res, err):
        if err:
            messagebox.showerror('Processing failed', str(err['value']))
        else:
            messagebox.showinfo(
                'Processing complete',
                'Invoices processed successfully.\n\n'
                f'PDFs: {res["pdf_count"]}\n'
                f'Items found: {res["items"]}\n'
                f'Updated: {res["updated"]}\n'
                f'Unchanged: {res["unchanged"]}\n'
                f'New/restored: {res["new"]}\n'
                f'Email: {res["email_status"]}\n\n'
                f'Workbook: {workbook}\n'
                f'Report: {res["report_path"]}'
            )
        root.destroy()

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    root.mainloop()


if __name__ == '__main__':
    supplier_name = sys.argv[1] if len(sys.argv) > 1 else 'Test'
    if supplier_name not in SUPPLIERS:
        raise SystemExit(f'Unknown supplier: {supplier_name}. Allowed: {supplier_names()}')
    pick_folder_and_run(supplier_name)
