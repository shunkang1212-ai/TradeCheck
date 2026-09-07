"""用户对本次会话提取值的修正。

只改内存里的副本，绝不写回原文件（任务书 5. / 6.）。
每次修正都记录原值、新值与来源单元格，并原样带进报告。
"""

from __future__ import annotations

import copy
from typing import Dict, List, Optional, Tuple

from .models import (
    FIELD_MISSING,
    FIELD_NEGATIVE,
    FIELD_NON_NUMERIC,
    FIELD_OK,
    Correction,
    DocumentExtraction,
    Field,
)
from .extract import field_kind
from .money import to_decimal


def clone_document(doc: DocumentExtraction) -> DocumentExtraction:
    return copy.deepcopy(doc)


def set_field_text(field: Field, new_text: str) -> None:
    """用新文本重刷一个字段的取值与状态。修正后的值视为人工确认值，不再是公式。"""
    text = (new_text or "").strip()
    field.text = text
    field.raw = text
    field.key = " ".join(text.split()).casefold()
    field.is_formula = False
    field.formula = None
    field.cached_available = False
    field.note = "用户修正值"

    if not text:
        field.state = FIELD_MISSING
        field.number = None
        return

    if field_kind(field.name) == "number":
        num = to_decimal(text)
        field.number = num
        if num is None:
            field.state = FIELD_NON_NUMERIC
        elif num < 0:
            field.state = FIELD_NEGATIVE
        else:
            field.state = FIELD_OK
    else:
        field.number = to_decimal(text) if _looks_numeric(text) else None
        field.state = FIELD_OK


def _looks_numeric(text: str) -> bool:
    return to_decimal(text) is not None


def _locate(doc: DocumentExtraction, target: str) -> Optional[Field]:
    """解析 target 并定位字段。

    支持：
      header.<field>        例如 header.currency
      row.<行号>.<field>    例如 row.9.qty
      total                 单据总额
    """
    parts = (target or "").split(".")
    if len(parts) == 2 and parts[0] == "header":
        return doc.header.get(parts[1])
    if len(parts) == 3 and parts[0] == "row":
        try:
            row_index = int(parts[1])
        except ValueError:
            return None
        for r in doc.rows:
            if r.row_index == row_index:
                return r.fields.get(parts[2])
        return None
    if len(parts) == 1 and parts[0] == "total":
        return doc.total
    return None


def apply_corrections(
    documents: Dict[str, DocumentExtraction], corrections: List[Correction]
) -> List[Correction]:
    """把修正应用到传入的（已克隆的）单据对象上，返回实际生效的修正记录。"""
    applied: List[Correction] = []
    for corr in corrections:
        if ":" not in corr.target:
            continue
        role, target = corr.target.split(":", 1)
        doc = documents.get(role)
        if doc is None:
            continue
        field = _locate(doc, target)
        if field is None:
            continue
        old_text = field.text
        if old_text == corr.new_text:
            continue

        cell = field.cell
        source = cell.text() if cell else doc.file_name
        set_field_text(field, corr.new_text)
        if target.startswith("header.") and field.present:
            # 人工补齐必需表头后，不保留已经失效的“字段缺失”提示。
            resolved_code = "missing-header-" + target.split(".", 1)[1]
            doc.issues = [issue for issue in doc.issues if issue.code != resolved_code]

        applied.append(
            Correction(
                target=corr.target,
                old_text=old_text,
                new_text=corr.new_text,
                source=source,
            )
        )
    return applied
