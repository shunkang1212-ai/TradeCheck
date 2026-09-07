"""把「读取 -> 提取 -> 检查 -> 报告」串起来的流程层。

界面（server.py）与命令行（cli.py）都只调用这里，业务逻辑不依赖界面。
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from . import report
from .corrections import apply_corrections, clone_document
from .excelio import WorkbookReadError, read_workbook
from .extract import extract_document
from .models import CheckResult, Correction, DocumentExtraction
from .rules import resolve_currency, run_checks
from .templates import TemplateError, load_template


class PipelineError(Exception):
    """面向用户的流程错误。"""


def extract_pair(
    invoice_path: str, packing_path: str, template_id: str,
    invoice_name: Optional[str] = None, packing_name: Optional[str] = None,
) -> Tuple[DocumentExtraction, DocumentExtraction, Dict[str, Any]]:
    template = load_template(template_id)
    try:
        inv_wb = read_workbook(invoice_path, invoice_name)
    except WorkbookReadError as exc:
        raise PipelineError(f"商业发票：{exc}")
    try:
        pac_wb = read_workbook(packing_path, packing_name)
    except WorkbookReadError as exc:
        raise PipelineError(f"装箱单：{exc}")

    invoice = extract_document(inv_wb, "invoice", template)
    packing = extract_document(pac_wb, "packing", template)
    return invoice, packing, template


def check_pair(
    invoice: DocumentExtraction,
    packing: DocumentExtraction,
    template: Dict[str, Any],
    corrections: Optional[List[Correction]] = None,
) -> Tuple[CheckResult, Dict[str, Any]]:
    """在克隆出来的副本上应用修正并执行检查，绝不改动传入的原始提取结果。"""
    inv = clone_document(invoice)
    pac = clone_document(packing)

    applied: List[Correction] = []
    if corrections:
        applied = apply_corrections({"invoice": inv, "packing": pac}, corrections)

    result = run_checks(inv, pac, template, applied)

    from .templates import money_config

    currency = resolve_currency(inv, money_config(template))
    extra = {
        "currency_text": currency.get("raw"),
        "currency_message": currency.get("message"),
        "currency_code": currency.get("code"),
        "currency_digits": currency.get("digits"),
    }
    return result, extra


def build_html_report(
    invoice: DocumentExtraction,
    packing: DocumentExtraction,
    result: CheckResult,
    template: Dict[str, Any],
    extra: Optional[Dict[str, Any]] = None,
    corrections: Optional[List[Correction]] = None,
) -> str:
    """生成报告 HTML。注意：报告展示的是修正后的数据，因此这里要用克隆+修正后的副本。"""
    inv = invoice
    pac = packing
    if corrections:
        inv = clone_document(invoice)
        pac = clone_document(packing)
        apply_corrections({"invoice": inv, "packing": pac}, corrections)
    return report.build_report(inv, pac, result, template, extra)


def render_report_file(
    out_path: str,
    invoice: DocumentExtraction,
    packing: DocumentExtraction,
    result: CheckResult,
    template: Dict[str, Any],
    extra: Optional[Dict[str, Any]] = None,
    corrections: Optional[List[Correction]] = None,
) -> str:
    html_text = build_html_report(invoice, packing, result, template, extra, corrections)
    return report.write_report(out_path, html_text)
