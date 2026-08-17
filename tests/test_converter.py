import io
import unittest
from datetime import datetime
from decimal import Decimal

from openpyxl import load_workbook
from pypdf import PdfWriter

from hdfc_converter import (
    PasswordRequiredError,
    Transaction,
    WrongPasswordError,
    create_workbook,
    parse_date,
    parse_money,
    unlock_pdf,
)


class ParsingTests(unittest.TestCase):
    def test_dates_and_common_ocr_damage(self):
        self.assertEqual(parse_date("10/05/25"), datetime(2025, 5, 10))
        self.assertEqual(parse_date("l0/05/25"), datetime(2025, 5, 10))
        self.assertEqual(parse_date("13l05/25"), datetime(2025, 5, 13))
        self.assertEqual(parse_date("t2105t25"), datetime(2025, 5, 12))
        self.assertIsNone(parse_date("not a date"))

    def test_money_and_split_paise(self):
        self.assertEqual(parse_money("40.000.00"), Decimal("40000.00"))
        self.assertEqual(parse_money("19.500 00"), Decimal("19500.00"))
        self.assertEqual(parse_money("0.00"), Decimal("0.00"))
        self.assertIsNone(parse_money(""))

    def test_password_handling(self):
        writer = PdfWriter()
        writer.add_blank_page(width=100, height=100)
        writer.encrypt("secret")
        encrypted = io.BytesIO()
        writer.write(encrypted)
        data = encrypted.getvalue()

        with self.assertRaises(PasswordRequiredError):
            unlock_pdf(data)
        with self.assertRaises(WrongPasswordError):
            unlock_pdf(data, "wrong")
        unlocked = unlock_pdf(data, "secret")
        self.assertFalse(__import__("pypdf").PdfReader(io.BytesIO(unlocked)).is_encrypted)

    def test_workbook_has_numeric_amounts(self):
        transaction = Transaction(
            datetime(2025, 5, 10), "Payment", "123", datetime(2025, 5, 10),
            Decimal("125.50"), Decimal("0"), Decimal("500.25"), 1,
        )
        data = create_workbook([transaction], {}, [])
        workbook = load_workbook(io.BytesIO(data), data_only=True)
        sheet = workbook["Transactions"]
        self.assertEqual(sheet["E2"].value, 125.5)
        self.assertEqual(sheet["G2"].value, 500.25)
        self.assertEqual(sheet.freeze_panes, "A2")


if __name__ == "__main__":
    unittest.main()
