"""金额与数量的十进制运算测试（R04 / R05 的基础）。

预期值全部手写硬编码，不调用被测函数生成。
"""

from __future__ import annotations

import os
import sys
import unittest
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tradecheck.money import parse_currency, quantize_amount, to_decimal


class TestToDecimal(unittest.TestCase):
    def test_plain_integers_and_decimals(self):
        self.assertEqual(to_decimal(1000), Decimal("1000"))
        self.assertEqual(to_decimal("1000"), Decimal("1000"))
        self.assertEqual(to_decimal("12.50"), Decimal("12.50"))

    def test_thousands_separator_and_spaces(self):
        self.assertEqual(to_decimal("1,234.50"), Decimal("1234.50"))
        self.assertEqual(to_decimal(" 1 234.50 "), Decimal("1234.50"))

    def test_currency_symbols(self):
        self.assertEqual(to_decimal("$1,000.00"), Decimal("1000.00"))
        self.assertEqual(to_decimal("¥3,150"), Decimal("3150"))

    def test_accounting_negative(self):
        self.assertEqual(to_decimal("(1,234.50)"), Decimal("-1234.50"))

    def test_non_numeric_returns_none_not_zero(self):
        for bad in ("N/A", "", "  ", "abc", None, True, False):
            self.assertIsNone(to_decimal(bad), f"{bad!r} 应返回 None，不能被当成 0")

    def test_float_does_not_leak_binary_noise(self):
        # 若直接 Decimal(0.1) 会得到 0.1000000000000000055511151231257827
        self.assertEqual(to_decimal(0.1), Decimal("0.1"))
        self.assertEqual(to_decimal(2.675), Decimal("2.675"))


class TestQuantize(unittest.TestCase):
    def test_half_up_classic_cases(self):
        self.assertEqual(quantize_amount(Decimal("0.125"), 2, "half_up"), Decimal("0.13"))
        self.assertEqual(quantize_amount(Decimal("2.675"), 2, "half_up"), Decimal("2.68"))
        self.assertEqual(quantize_amount(Decimal("0.124"), 2, "half_up"), Decimal("0.12"))

    def test_half_up_is_not_half_even(self):
        # 0.125 用 HALF_EVEN 会得到 0.12，用 HALF_UP 得到 0.13
        self.assertEqual(quantize_amount(Decimal("0.125"), 2, "half_up"), Decimal("0.13"))
        self.assertEqual(quantize_amount(Decimal("0.125"), 2, "half_even"), Decimal("0.12"))

    def test_zero_digit_currency(self):
        self.assertEqual(quantize_amount(Decimal("1234.56"), 0, "half_up"), Decimal("1235"))

    def test_repeatable(self):
        for _ in range(50):
            self.assertEqual(quantize_amount(Decimal("7.035"), 2, "half_up"), Decimal("7.04"))


class TestParseCurrency(unittest.TestCase):
    def test_iso_code(self):
        code, digits, _ = parse_currency("USD")
        self.assertEqual((code, digits), ("USD", 2))

    def test_chinese_alias(self):
        code, digits, _ = parse_currency("人民币 CNY")
        self.assertEqual((code, digits), ("CNY", 2))
        code, digits, _ = parse_currency("美元")
        self.assertEqual((code, digits), ("USD", 2))

    def test_ambiguous_symbol_requires_confirmation(self):
        code, digits, _ = parse_currency("$")
        self.assertEqual((code, digits), (None, None))

    def test_zero_decimal_currency_is_not_assumed_two(self):
        code, digits, _ = parse_currency("JPY")
        self.assertEqual((code, digits), ("JPY", 0))

    def test_unknown_currency_is_not_silently_accepted(self):
        code, digits, msg = parse_currency("XXX")
        self.assertIsNone(code)
        self.assertIsNone(digits)
        self.assertIn("人工确认", msg)

    def test_empty(self):
        code, digits, msg = parse_currency("")
        self.assertIsNone(code)
        self.assertIsNone(digits)


if __name__ == "__main__":
    unittest.main(verbosity=2)
