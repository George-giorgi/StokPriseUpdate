# StokPriseUpdate — Chadwicks Invoice Price Update

## What it does

This Python script reads Chadwicks invoice PDFs from `Full_Invoices` and updates the Excel price file.

**Workflow:**

New invoices → `Full_Invoices` → Run script → Excel updated → Email report sent

The script can:
- Read one or many Chadwicks invoices.
- Find Material IDs.
- Update prices when the invoice price differs from Excel.
- Keep previous price and price history.
- Calculate Difference and Change %.
- Restore a Material ID if its Excel row was manually deleted but it appears on an invoice.
- Add new Material IDs found on invoices.
- Record the latest invoice information.
- Report price changes greater than 10%.
- Send the report by Gmail.

---

## Important files and folders

### `Full_Invoices\`
Temporary input folder.

Put **only the new invoice PDFs** here.

After successful processing, the folder can be emptied and reused for the next batch.

### `Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx`
The Excel workbook/database.

**Keep this file. Do not delete it unless you intentionally want to start again.**

### `chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py`
The current working script.

This is the script to use for normal processing.

### `CHAD_Invoices\`
Older/separate invoice folder. The normal command uses `Full_Invoices`.

### `.gitignore`
Helps prevent private invoice files from being uploaded to GitHub.

---

# How to run

## 1. Open Command Prompt

Open **CMD**.

## 2. Go to the project folder

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
```

## 3. Put new invoices into `Full_Invoices`

Example:

```text
Full_Invoices\
    SX236100.pdf
    SX236101.pdf
    SX236102.pdf
```

You can process one invoice or many at the same time.

## 4. Run the script

```cmd
python Chadwicks\chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py "Chadwicks/Full_Invoices" "Chadwicks/Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx"
```

## 5. Check the result

CMD should show a summary similar to:

```text
PDFs found: ...
Items detected: ...
Prices updated: ...
Prices unchanged: ...
New materials: ...
Materials restored: ...
Needs review: ...
Errors: ...
Email report: SENT
```

Then check the Excel workbook and the email report.

---

# Normal day-to-day workflow

1. Empty old processed PDFs from `Full_Invoices`.
2. Drop the new invoice PDF/PDFs into `Full_Invoices`.
3. Open CMD.
4. Run:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
```

5. Run:

```cmd
python Chadwicks\chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py "Chadwicks/Full_Invoices" "Chadwicks/Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx"
```

6. Check Excel and the email report.
7. After confirming successful processing, empty `Full_Invoices` ready for the next batch.

---

# Price update logic

**Material ID is the key.**

Example:

```text
Material ID: 29011
Excel price: €5.50
Invoice price: €8.29
```

The script updates the current price and records the previous price/history.

The report can show:

```text
29011 | GYPROC WALLBOARD... | €5.50 → €8.29 | +50.73%
```

---

# Price changes over 10%

The email contains:

```text
Materials with price changes over 10%:
```

Each Material ID is shown once using the latest >10% change processed in that run.

---

# Restoring deleted materials

If a Material ID is manually deleted from Excel but appears on a new invoice, the script can add it back.

The restored material can include:
- Material ID
- Description
- Unit
- Supplier
- Invoice price
- Last Invoice
- Last Updated

---

# New materials

If a Material ID appears on an invoice but does not exist in Excel, the script can add it as a new material.

---

# Email report

The script sends a report through Gmail containing:
- Invoices processed
- Materials checked
- Prices updated
- Prices unchanged
- New materials
- Restored materials
- Review/errors
- Materials with price changes over 10%

The CSV report is attached to the email.

The Excel workbook is not attached automatically.

---

# Important safety rules

### Do not upload real invoices to GitHub

Real invoices may contain private company/customer/account/order/pricing information.

Keep them in:

```text
Full_Invoices
```

and make sure the folder is covered by `.gitignore`.

### Do not delete the Excel workbook

It contains the accumulated material and price information.

### Do not share the Gmail App Password

The App Password is stored locally using Windows Credential Manager.

Never put it inside the Python script or GitHub.

---

# If Python does not run

Check Python:

```cmd
python --version
```

If required packages are missing:

```cmd
pip install pdfplumber openpyxl keyring
```

---

# If email does not send

The script uses Gmail SMTP.

It uses:
- `CHADWICKS_GMAIL_USER`
- `CHADWICKS_BOSS_EMAIL`
- Gmail App Password stored in Windows Credential Manager.

Do not use your normal Gmail password in the Python script.

---

# Quick command

For normal use, these are the two commands:

```cmd
cd "C:\Users\MauriceSweeney(Brusn\StokPriseUpdate\StokPriseUpdate"
```

```cmd
python chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py "Full_Invoices" "Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx"
```

## Project structure

```text
StokPriseUpdate
│
├── .gitignore
├── README_StokPriseUpdate.md
├── Chadwicks_FULL_INVOICES_TEST_NO_MATERIAL.xlsx
├── chadwicks_invoice_import_FINAL4_EMAIL_OVER10_LATEST_ONLY.py
├── CHAD_Invoices\
└── Full_Invoices\
    └── NEW INVOICE PDFs GO HERE
```

## Simple rule

**Drop new invoices → run script → check Excel/email → empty `Full_Invoices` → ready for the next batch.**
