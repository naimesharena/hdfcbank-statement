# Scanned Bank Statement to Excel

A local Streamlit app that converts scanned bank statements into Excel using OCR. It supports a file-selection control, password-protected PDFs, page ranges, and PDF/JPG/PNG/TIFF input.

## What the workbook contains

- **Transactions** — date, narration, reference, value date, withdrawal, deposit, balance, page, and the source OCR line
- **OCR Text** — recognized text per page plus average OCR confidence
- **Conversion Info** — source and conversion summary

> OCR can misread financial data. Always compare dates, amounts, totals, and closing balances with the original statement.

## Run locally

### 1. Install Tesseract

Ubuntu/Debian:

```bash
sudo apt-get update
sudo apt-get install tesseract-ocr tesseract-ocr-eng
```

macOS:

```bash
brew install tesseract
```

Windows: install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) and ensure `tesseract.exe` is on `PATH`.

### 2. Install and run the app

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Open the URL shown by Streamlit, select the statement, enter its PDF password if needed, choose the page range, and click **Convert to Excel**.

### Windows shortcuts

You can also double-click **`run_app.bat`**, or run the Python file from an IDE:

```powershell
python app.py
```

`app.py` detects a normal Python launch and automatically restarts itself using Streamlit. This avoids the `missing ScriptRunContext` warnings produced when a Streamlit script is run in bare mode.

## Password handling and privacy

The upload is processed in memory. A supplied password is passed directly to PyMuPDF to unlock the document; it is not written to disk or included in the output workbook. When deploying publicly, remember that processing occurs on the server running Streamlit.

## Deploy on Streamlit Community Cloud

Push these files to GitHub and create a Streamlit app with `app.py` as the entry point. `packages.txt` asks the platform to install Tesseract.

## Parser notes

The parser targets common HDFC layouts where a transaction row starts with a date and ends with transaction amount/balance columns. Wrapped descriptions are joined. When OCR collapses blank debit/credit columns, direction is inferred from consecutive balances. The original OCR line is always retained so uncertain rows are auditable.
