import statement_converter
from statement_converter import OCRPage, parse_transactions


def test_parses_hdfc_rows_and_infers_direction():
    pages = [OCRPage(1, """Date Narration Chq./Ref.No. Value Dt Withdrawal Deposit Closing Balance
01/08/2026 OPENING BALANCE 01/08/2026 0.00 10,000.00
02/08/2026 UPI-SHOP 123456789012 02/08/2026 250.00 9,750.00
03/08/2026 SALARY CREDIT 03/08/2026 5,000.00 14,750.00""", 90)]
    rows = parse_transactions(pages)
    assert len(rows) == 3
    assert rows[1].reference == "123456789012"
    assert rows[1].withdrawal == 250
    assert rows[1].deposit is None
    assert rows[2].deposit == 5000
    assert rows[2].balance == 14750


def test_joins_wrapped_narration():
    pages = [OCRPage(2, """04-08-26 NEFT TRANSFER 04-08-26 100.00 14,650.00
TO EXAMPLE BENEFICIARY
05-08-26 NEXT ITEM 05-08-26 50.00 14,600.00""", 80)]
    rows = parse_transactions(pages)
    assert "EXAMPLE BENEFICIARY" in rows[0].narration
    assert rows[0].page == 2


def test_configures_tesseract_from_environment(tmp_path, monkeypatch):
    executable = tmp_path / "tesseract.exe"
    executable.touch()
    monkeypatch.setenv("TESSERACT_CMD", str(executable))
    monkeypatch.setattr(statement_converter.shutil, "which", lambda _: None)

    statement_converter.check_tesseract()

    assert statement_converter.pytesseract.pytesseract.tesseract_cmd == str(executable)
