"""金额与数量的十进制运算（规则 R04 / R05 的基础）。

严格不使用二进制浮点做金额比较：所有金额与数量都以 decimal.Decimal 承载。
float 输入一律经 repr() 转成十进制字符串再构造 Decimal，避免 0.1 之类的二进制误差。
"""

from __future__ import annotations

import decimal
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Optional, Tuple

#: 已知币种的小数位数。不在表里的币种一律视为「未知」，相关检查降级，不假定两位小数。
KNOWN_CURRENCY_DIGITS = {
    "USD": 2, "EUR": 2, "GBP": 2, "CHF": 2, "CAD": 2, "AUD": 2, "NZD": 2,
    "CNY": 2, "HKD": 2, "TWD": 2, "SGD": 2, "MYR": 2, "THB": 2, "INR": 2,
    "RUB": 2, "BRL": 2, "MXN": 2, "ZAR": 2, "TRY": 2, "AED": 2, "SAR": 2,
    "PLN": 2, "SEK": 2, "NOK": 2, "DKK": 2, "CZK": 2, "IDR": 2, "PHP": 2,
    "VND": 0, "JPY": 0, "KRW": 0, "CLP": 0, "ISK": 0,
}

#: 中文币种名 -> ISO 代码
CURRENCY_ALIASES = {
    "美元": "USD", "美金": "USD", "美圆": "USD",
    "人民币": "CNY", "元": "CNY", "RMB": "CNY", "CNY¥": "CNY",
    "欧元": "EUR", "英镑": "GBP", "日元": "JPY", "日圆": "JPY",
    "港币": "HKD", "港元": "HKD", "韩元": "KRW", "新台币": "TWD",
    "新加坡元": "SGD", "澳元": "AUD", "加元": "CAD", "瑞士法郎": "CHF",
}

_CURRENCY_RE = re.compile(r"\b([A-Za-z]{3})\b")
_SYMBOL_MAP = {"$": "USD", "¥": "CNY", "￥": "CNY", "€": "EUR", "£": "GBP"}

#: 数量/金额字符串里需要剥离的噪音字符
_NOISE_RE = re.compile(r"[,\s 　_']")


def to_decimal(value) -> Optional[Decimal]:
    """把单元格原始值转成 Decimal；无法解析时返回 None（绝不抛异常、绝不默认成 0）。"""
    if value is None:
        return None
    if isinstance(value, bool):  # Excel TRUE/FALSE 不是数字
        return None
    if isinstance(value, Decimal):
        return value if value.is_finite() else None
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return None
        # 经 repr 转十进制，避免 Decimal(0.1) 得到 0.1000000000000000055511151231257827
        try:
            return Decimal(repr(value))
        except InvalidOperation:
            return None

    text = str(value).strip()
    if not text:
        return None

    negative = False
    if text.startswith("(") and text.endswith(")"):  # 会计负数写法 (1,234.00)
        negative = True
        text = text[1:-1]

    # 逗号仅作为三位分组符；“1,25”不能被静默读成 125。
    comma_text = re.sub(r"[\s 　_']", "", text).strip("$¥￥€£")
    if "," in comma_text and not re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", comma_text):
        return None
    text = _NOISE_RE.sub("", text)
    # 剥掉前后货币符号
    text = text.lstrip("$¥￥€£").rstrip("$¥￥€£")
    if not text:
        return None

    try:
        d = Decimal(text)
    except InvalidOperation:
        return None
    if not d.is_finite():
        return None
    return -d if negative else d


def parse_currency(text: str) -> Tuple[Optional[str], Optional[int], str]:
    """从币种字段文本解析 ISO 代码与小数位数。

    返回 (code, digits, message)。解析不出来时 code 为 None，message 说明原因，
    调用方必须把相关检查降级为「待确认 / 未检查」，不能假定两位小数。
    """
    raw = (text or "").strip()
    if not raw:
        return None, None, "币种字段为空"

    # 先消费完整别名，避免“元”抢先匹配“欧元/日元/港元”。
    remainder = raw.upper()
    codes = set()
    for alias, code in sorted(CURRENCY_ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        if alias == "元":
            continue  # 单独“元”的币种不确定。
        if alias.upper() in remainder:
            codes.add(code)
            remainder = remainder.replace(alias.upper(), " ")
    for match in _CURRENCY_RE.finditer(remainder):
        codes.add(match.group(1))
    remainder = _CURRENCY_RE.sub("", remainder)
    for sym, code in {"€": "EUR", "£": "GBP"}.items():
        if sym in remainder:
            codes.add(code)
    # $、¥ 等单独出现时含义不唯一；必须有明确代码或完整币种名。
    residual = re.sub(r"[\s$¥￥€£/()（）,，:：-]", "", remainder)
    if len(codes) != 1 or residual:
        return None, None, "币种缺失、歧义或包含多个币种，请填写单一明确币种代码并人工确认"
    code = next(iter(codes))
    digits = KNOWN_CURRENCY_DIGITS.get(code)
    if digits is None:
        return None, None, f"币种代码 {code} 的小数位数规则未收录，需人工确认"
    return code, digits, f"识别为 {code}"


_ROUNDING_MODES = {
    "half_up": ROUND_HALF_UP,
    "half_even": decimal.ROUND_HALF_EVEN,
    "down": decimal.ROUND_DOWN,
    "up": decimal.ROUND_UP,
}


def rounding_mode(name: str):
    return _ROUNDING_MODES.get(str(name).strip().casefold(), ROUND_HALF_UP)


def quantize_amount(amount: Decimal, digits: int, mode_name: str = "half_up") -> Decimal:
    exp = Decimal(1).scaleb(-int(digits))
    return amount.quantize(exp, rounding=rounding_mode(mode_name))


def format_decimal(value: Optional[Decimal], digits: Optional[int] = None) -> str:
    if value is None:
        return "—"
    if digits is not None:
        try:
            return f"{quantize_amount(value, digits):,.{int(digits)}f}"
        except Exception:
            pass
    return f"{value:,.10g}"


def within_tolerance(a: Decimal, b: Decimal, tolerance: Decimal) -> bool:
    return abs(a - b) <= tolerance
