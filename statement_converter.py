"""OCR and Excel export helpers for scanned bank statements.

The module has no Streamlit dependency, which makes it easy to test or reuse from
another UI. PDF pages are rendered in memory and uploaded statements are never
written to disk.
"""
from __future__ import annotations

import io
import os
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

import pymupdf
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

# Dates also tolerate common scanner substitutions (O -> 0, I/l -> 1).
_DATE_DIGIT = r"0-9OoIl"
DATE_PATTERN = re.compile(
    rf"(?<![{_DATE_DIGIT}])([{_DATE_DIGIT}]{{1,2}}[./-]"
    rf"[{_DATE_DIGIT}]{{1,2}}[./-][{_DATE_DIGIT}]{{2,4}})(?![{_DATE_DIGIT}])"
)
# Indian statements normally use 1,234.56, but Tesseract frequently reads the
# decimal point as a comma (for example 390.00 -> 390,00). Accept both forms.
_DIGIT = r"\dOoIl"
AMOUNT_PATTERN = re.compile(
    rf"(?<!\w)(?:₹\s*)?([+-]?(?:[{_DIGIT}][{_DIGIT},]*\.[{_DIGIT}]{{2}}|"
    rf"[{_DIGIT}]+,[{_DIGIT}]{{2}}))(?:\s*(?:CR|DR))?",
    re.I,
)


class StatementError(Exception):
    """A message that can safely be shown in the UI."""


@dataclass
class OCRPage:
    page: int
    text: str
    confidence: float


@dataclass
class Transaction:
    date: str
    narration: str
    reference: str
    value_date: str
    withdrawal: float | None
    deposit: float | None
    balance: float | None
    page: int
    raw_text: str


def find_tesseract() -> str | None:
    """Locate Tesseract on PATH or in its common Windows install folders."""
    configured = os.environ.get("TESSERACT_CMD", "").strip().strip('"')
    candidates = [configured] if configured else []

    on_path = shutil.which("tesseract") or shutil.which("tesseract.exe")
    if on_path:
        candidates.append(on_path)

    # The Windows installer does not always add Tesseract to PATH.
    candidates.extend([
        str(Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Tesseract-OCR" / "tesseract.exe"),
        str(Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Tesseract-OCR" / "tesseract.exe"),
        str(Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe"),
    ])
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(Path(candidate))
    return None


def check_tesseract() -> None:
    """Configure pytesseract or raise a friendly installation error."""
    executable = find_tesseract()
    if not executable:
        raise StatementError(
            "Tesseract OCR was not found. On Windows it is normally installed at "
            "C:\\Program Files\\Tesseract-OCR\\tesseract.exe. If yours is elsewhere, "
            "set the TESSERACT_CMD environment variable to its full path and restart the app."
        )
    pytesseract.pytesseract.tesseract_cmd = executable


def open_document(data: bytes, filename: str, password: str = ""):
    """Open an uploaded PDF/image and authenticate encrypted PDFs."""
    suffix = filename.rsplit(".", 1)[-1].lower() if "." in filename else "pdf"
    try:
        if suffix == "pdf":
            doc = pymupdf.open(stream=data, filetype="pdf")
            if doc.needs_pass and not password:
                doc.close()
                raise StatementError("This PDF is password protected. Enter its password.")
            if doc.needs_pass and not doc.authenticate(password):
                doc.close()
                raise StatementError("The PDF password is incorrect.")
            return doc

        image = Image.open(io.BytesIO(data))
        frames: list[bytes] = []
        for index in range(getattr(image, "n_frames", 1)):
            image.seek(index)
            frame = image.convert("RGB")
            output = io.BytesIO()
            frame.save(output, format="PDF", resolution=250)
            frames.append(output.getvalue())
        merged = pymupdf.open()
        for frame_pdf in frames:
            source = pymupdf.open(stream=frame_pdf, filetype="pdf")
            merged.insert_pdf(source)
            source.close()
        return merged
    except StatementError:
        raise
    except Exception as exc:
        raise StatementError(f"Could not open the selected file: {exc}") from exc


def _prepare_image(page: pymupdf.Page, dpi: int) -> Image.Image:
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY, alpha=False)
    image = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    # Autocontrast removes a grey scanner background without destroying thin text.
    image = ImageOps.autocontrast(image, cutoff=1)
    image = ImageEnhance.Contrast(image).enhance(1.35)
    return image.filter(ImageFilter.SHARPEN)


def ocr_document(
    doc,
    first_page: int = 1,
    last_page: int | None = None,
    dpi: int = 250,
    language: str = "eng",
    progress: Callable[[int, int], None] | None = None,
) -> list[OCRPage]:
    """OCR a page range. Page numbers accepted and returned are one-based."""
    check_tesseract()
    last_page = min(last_page or doc.page_count, doc.page_count)
    if first_page < 1 or first_page > last_page:
        raise StatementError("Select a valid page range.")

    pages: list[OCRPage] = []
    total = last_page - first_page + 1
    config = "--oem 3 --psm 6 preserve_interword_spaces=1"
    for done, page_number in enumerate(range(first_page, last_page + 1), start=1):
        image = _prepare_image(doc.load_page(page_number - 1), dpi)
        data = pytesseract.image_to_data(
            image, lang=language, config=config, output_type=pytesseract.Output.DICT
        )
        # Reconstruct lines from Tesseract's layout data, retaining useful spacing.
        grouped: dict[tuple[int, int, int], list[tuple[int, str]]] = {}
        confidences: list[float] = []
        for i, word in enumerate(data["text"]):
            word = word.strip()
            if not word:
                continue
            key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            grouped.setdefault(key, []).append((int(data["left"][i]), word))
            try:
                confidence = float(data["conf"][i])
                if confidence >= 0:
                    confidences.append(confidence)
            except (TypeError, ValueError):
                pass
        lines = [" ".join(word for _, word in sorted(words)) for words in grouped.values()]
        pages.append(
            OCRPage(
                page=page_number,
                text="\n".join(lines),
                confidence=round(sum(confidences) / len(confidences), 1) if confidences else 0,
            )
        )
        if progress:
            progress(done, total)
    return pages


def _amount(value: str) -> float | None:
    cleaned = value.replace("₹", "").replace(" ", "")
    # Common OCR substitutions inside numeric fields.
    cleaned = cleaned.translate(str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1"}))
    if "." in cleaned:
        # A period is the decimal mark, so any commas are thousands separators.
        cleaned = cleaned.replace(",", "")
    elif cleaned.count(",") == 1:
        # OCR often changes a decimal point to a comma: 390.00 -> 390,00.
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _normalise_date(value: str) -> str:
    value = value.translate(str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1"}))
    value = value.replace(".", "/").replace("-", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(value, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return value


def parse_transactions(pages: Iterable[OCRPage]) -> list[Transaction]:
    """Parse common HDFC-style rows and retain raw OCR for manual checking.

    A row begins with a transaction date. Wrapped narration lines are appended to
    that row. The last amount is treated as balance and the prior amount as the
    transaction value. Deposit/withdrawal is inferred from consecutive balances.
    """
    candidates: list[dict] = []
    current: dict | None = None
    for page in pages:
        for source_line in page.text.splitlines():
            line = " ".join(source_line.split())
            date_matches = list(DATE_PATTERN.finditer(line))
            starts_row = bool(date_matches and date_matches[0].start() <= 4)
            if starts_row:
                if current:
                    candidates.append(current)
                first = date_matches[0]
                second = date_matches[1] if len(date_matches) > 1 else None
                amount_start = second.end() if second else first.end()
                amounts = [_amount(m.group(1)) for m in AMOUNT_PATTERN.finditer(line[amount_start:])]
                amounts = [amount for amount in amounts if amount is not None]
                middle_end = second.start() if second else (
                    next((m.start() + amount_start for m in AMOUNT_PATTERN.finditer(line[amount_start:])), len(line))
                )
                middle = line[first.end():middle_end].strip(" |-:")
                # A cheque/reference number is usually the final long token that
                # contains at least one digit (avoids treating narration as a ref).
                ref_candidates = re.findall(r"(?:^|\s)([A-Z0-9/-]{6,})(?=\s|$)", middle, re.I)
                ref_candidates = [value for value in ref_candidates if any(ch.isdigit() for ch in value)]
                reference = ref_candidates[-1] if ref_candidates else ""
                narration = middle.replace(reference, "", 1).strip(" |-:") if reference else middle
                current = {
                    "date": _normalise_date(first.group(1)),
                    "value_date": _normalise_date(second.group(1)) if second else "",
                    "narration": narration,
                    "reference": reference,
                    "amounts": amounts,
                    "page": page.page,
                    "raw": [line],
                }
            elif current and line and not re.search(r"opening balance|closing balance|statement summary", line, re.I):
                # Wrapped narrations generally occur before the next dated row.
                current["raw"].append(line)
                if len(current["raw"]) <= 3:
                    current["narration"] = (current["narration"] + " " + line).strip()
        if current:
            candidates.append(current)
            current = None

    transactions: list[Transaction] = []
    for item in candidates:
        amounts = item["amounts"]
        balance = amounts[-1] if amounts else None
        transaction_amount = amounts[-2] if len(amounts) >= 2 else None
        withdrawal = deposit = None
        # If OCR preserved all three financial columns, use them directly.
        if len(amounts) >= 3:
            withdrawal = amounts[-3] or None
            deposit = amounts[-2] or None
        transactions.append(Transaction(
            date=item["date"], narration=item["narration"], reference=item["reference"],
            value_date=item["value_date"], withdrawal=withdrawal, deposit=deposit,
            balance=balance, page=item["page"], raw_text=" | ".join(item["raw"]),
        ))
        # Keep the unresolved amount temporarily on the object.
        setattr(transactions[-1], "_transaction_amount", transaction_amount)

    # Infer debit/credit from adjacent balances when blank columns collapsed in OCR.
    for index, row in enumerate(transactions):
        if row.withdrawal is not None or row.deposit is not None:
            continue
        amount = getattr(row, "_transaction_amount", None)
        # With only "0.00 + balance", OCR probably missed the non-zero debit or
        # credit column. Leave both blank instead of reporting a false zero.
        if amount is None or amount == 0:
            continue
        previous = transactions[index - 1] if index else None
        if previous and previous.balance is not None and row.balance is not None:
            delta = row.balance - previous.balance
            if abs(abs(delta) - amount) <= max(1.0, amount * 0.01):
                if delta < 0:
                    row.withdrawal = amount
                else:
                    row.deposit = amount
                continue
        # Direction remains unknown on first/unclear rows; retain amount as withdrawal
        # and flag it in raw output rather than silently dropping it.
        row.withdrawal = amount
    return transactions


def create_workbook(transactions: list[Transaction], pages: list[OCRPage], source_name: str) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Transactions"
    headers = ["Date", "Narration", "Reference / Cheque No.", "Value Date", "Withdrawal", "Deposit", "Balance", "Source Page", "Raw OCR"]
    sheet.append(headers)
    for row in transactions:
        sheet.append([row.date, row.narration, row.reference, row.value_date, row.withdrawal, row.deposit, row.balance, row.page, row.raw_text])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column in (5, 6, 7):
        for cell in sheet.iter_cols(min_col=column, max_col=column, min_row=2):
            cell[0].number_format = '#,##0.00;[Red]-#,##0.00'
    widths = [14, 55, 24, 14, 16, 16, 16, 12, 70]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
        row[8].alignment = Alignment(wrap_text=True, vertical="top")

    raw = workbook.create_sheet("OCR Text")
    raw.append(["Page", "Average OCR confidence", "Recognized text"])
    for page in pages:
        raw.append([page.page, page.confidence, page.text])
    for cell in raw[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
    raw.column_dimensions["A"].width = 10
    raw.column_dimensions["B"].width = 24
    raw.column_dimensions["C"].width = 120
    for row in raw.iter_rows(min_row=2):
        row[2].alignment = Alignment(wrap_text=True, vertical="top")

    info = workbook.create_sheet("Conversion Info")
    info.append(["Source file", source_name])
    info.append(["Pages processed", len(pages)])
    info.append(["Transactions detected", len(transactions)])
    info.append(["Important", "OCR can make mistakes. Compare totals, dates and balances with the original statement before relying on this workbook."])
    info.column_dimensions["A"].width = 24
    info.column_dimensions["B"].width = 110
    info["A1"].font = Font(bold=True)
    info["B4"].alignment = Alignment(wrap_text=True)

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()
