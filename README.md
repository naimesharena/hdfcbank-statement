# HDFC Bank Statement to Excel

A local web application that converts the standard HDFC Bank PDF statement layout into a formatted `.xlsx` workbook.

## Features

- PDF file-selection/drop menu
- Password field for encrypted PDF statements
- Extracts date, narration, reference number, value date, withdrawal, deposit, and closing balance
- Creates a filterable Excel table with Indian bank amounts preserved as numbers
- Adds statement metadata and conversion statistics on a second sheet
- Marks OCR fields that need manual review
- Processes the PDF and password in memory; the app does not save either one

## Single-file desktop version

`hdfc_statement_to_excel.py` is completely self-contained application code. Running it immediately opens the operating system's PDF file-selection window. If the selected PDF is encrypted, it then displays a masked password prompt, followed by an Excel save dialog.

Python 3.10 or newer is recommended. Install the libraries once:

```bash
pip install -r requirements.txt
```

The added OCR libraries render each statement page at high resolution and read the scan again instead of trusting the PDF's often-corrupted hidden text. Processing a long statement can take several minutes, but dates and amounts are substantially more accurate and all processing remains local.

### Poppler and Tesseract on Windows

Install both desktop programs, then update the configuration near the top of `hdfc_statement_to_excel.py` if your locations differ:

```python
OCR_ENGINE = "rapidocr"  # recommended for this HDFC scan
POPPLER_PATH = r"C:\poppler\Library\bin"
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
OCR_DPI = 300
```

`rapidocr` is the verified default for the supplied statement. Set `OCR_ENGINE = "tesseract"` only to force Poppler + Tesseract, or use `"auto"` to try RapidOCR first and Tesseract only as a fallback.

Then run the single file:

```bash
python hdfc_statement_to_excel.py
```

On some Ubuntu/Debian installations, the desktop dialog library must first be installed with `sudo apt install python3-tk`. Tkinter is already included with normal Python installations on Windows and macOS.

## Browser version

A Streamlit browser interface is also available:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Select the statement, enter a password only when the PDF is protected, and click **Convert to Excel**.

## Supported files

The parser targets HDFC's standard statement-of-account PDF layout. It works with normal text PDFs and scanned PDFs that already contain an OCR text layer. A raw image-only scan must be OCR-processed before conversion.

PDF OCR is not always exact. The generated workbook includes the source PDF page and a **Review Notes** column, so uncertain dates or balances can be checked against the statement.

## Use from Python

```python
from pathlib import Path
from hdfc_converter import convert_statement

pdf = Path("statement.pdf").read_bytes()
result = convert_statement(pdf, password="optional-password")
Path("statement.xlsx").write_bytes(result.workbook)
```

The password can be omitted for an unprotected statement.
