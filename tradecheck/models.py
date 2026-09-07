"""TradeCheck 数据模型。

设计要点：
- 每个取值都同时保留「原始字符串」和「规范化值」两份数据（任务书 3.）。
- 型号等标识类字段的前导零必须保留，因此内部一律以字符串承载，不做数值转换。
- 缺失值一律是 None / 空，不转换成 0（任务书 3.）。
- 公式单元格单独记录，缓存值未经重新计算，标记 cached_available=False（任务书 R06）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Dict, List, Optional

# ---------------------------------------------------------------- 状态常量

#: 输入及规则具备检查条件，检查完成且无差异
STATUS_OK = "未发现差异"
#: 有可复现的比较或计算依据
STATUS_DIFF = "发现差异"
#: 识别、匹配或业务含义不明确
STATUS_TODO = "待确认"
#: 超出支持范围或缺少检查条件
STATUS_SKIP = "未检查"

STATUS_ORDER = {STATUS_DIFF: 0, STATUS_TODO: 1, STATUS_SKIP: 2, STATUS_OK: 3}

#: 字段取值状态
FIELD_OK = "ok"
FIELD_MISSING = "missing"
FIELD_NON_NUMERIC = "non_numeric"
FIELD_NEGATIVE = "negative"
FIELD_UNVERIFIED = "unverified"  # 公式缓存未验证


def worst_status(statuses: List[str]) -> str:
    """取一组状态中最严重的一个（差异 > 待确认 > 未检查 > 无差异）。"""
    if not statuses:
        return STATUS_OK
    return sorted(statuses, key=lambda s: STATUS_ORDER.get(s, 9))[0]


# ---------------------------------------------------------------- 单元格引用


@dataclass(frozen=True)
class CellRef:
    """一个取值在原始文件中的位置，用于报告中的证据追溯。"""

    file_name: str
    sheet: str
    coord: str  # 例如 "C9"

    def text(self) -> str:
        return f"{self.file_name} · {self.sheet}!{self.coord}"

    def to_dict(self) -> Dict[str, str]:
        return {"file_name": self.file_name, "sheet": self.sheet, "coord": self.coord}


# ---------------------------------------------------------------- 字段取值


@dataclass
class Field:
    """一个被提取出来的字段。

    text  : 原始字符串，保留前导零与原文空格，用于展示与型号匹配。
    number: 解析后的十进制数值，仅在需要算术时使用；无法解析时为 None。
    key   : 规范化匹配键（型号/规格比对用），不做数值化。
    """

    name: str = ""
    text: str = ""
    raw: Any = None
    number: Optional[Decimal] = None
    key: str = ""
    cell: Optional[CellRef] = None
    state: str = FIELD_MISSING
    is_formula: bool = False
    formula: Optional[str] = None
    cached_available: bool = False
    note: str = ""

    @property
    def present(self) -> bool:
        return self.state not in (FIELD_MISSING,) and self.text.strip() != ""

    @property
    def usable_number(self) -> Optional[Decimal]:
        """可用于计算的数值；公式缓存未验证时返回 None，强制相关检查降级。"""
        if self.state in (FIELD_MISSING, FIELD_NON_NUMERIC, FIELD_UNVERIFIED):
            return None
        return self.number

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "text": self.text,
            "number": None if self.number is None else str(self.number),
            "key": self.key,
            "cell": self.cell.to_dict() if self.cell else None,
            "state": self.state,
            "is_formula": self.is_formula,
            "formula": self.formula,
            "cached_available": self.cached_available,
            "note": self.note,
        }


# ---------------------------------------------------------------- 明细行 / 单据


@dataclass
class DetailRow:
    row_index: int  # Excel 行号，1 起
    fields: Dict[str, Field] = field(default_factory=dict)
    flags: List[str] = field(default_factory=list)  # 例如 "total_row"

    def get(self, name: str) -> Field:
        return self.fields.get(name) or Field(name=name)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "row_index": self.row_index,
            "fields": {k: v.to_dict() for k, v in self.fields.items()},
            "flags": list(self.flags),
        }


@dataclass
class Issue:
    """读取/提取阶段发现的问题，尚未进入规则判定。"""

    code: str
    message: str
    cell: Optional[CellRef] = None
    severity: str = STATUS_TODO  # 待确认 / 未检查

    def to_dict(self) -> Dict[str, Any]:
        return {
            "code": self.code,
            "message": self.message,
            "cell": self.cell.to_dict() if self.cell else None,
            "severity": self.severity,
        }


@dataclass
class RowSnapshot:
    """明细区某行的快照，供「额外费用结构」等需要看整行内容的判断使用。

    只记录标签文本与同行数值及其单元格位置，不记录无关正文。
    """

    row_index: int
    label: str = ""
    label_cell: Optional[CellRef] = None
    numbers: List[Dict[str, Any]] = field(default_factory=list)  # {text, number, cell}

    def to_dict(self) -> Dict[str, Any]:
        return {
            "row_index": self.row_index,
            "label": self.label,
            "label_cell": self.label_cell.to_dict() if self.label_cell else None,
            "numbers": [
                {
                    "text": n.get("text", ""),
                    "number": None if n.get("number") is None else str(n["number"]),
                    "cell": n["cell"].to_dict() if n.get("cell") else None,
                }
                for n in self.numbers
            ],
        }


@dataclass
class DocumentExtraction:
    """一份单据（商业发票或装箱单）的提取结果。"""

    role: str  # "invoice" | "packing"
    role_label: str
    file_name: str
    sha256: str
    template_id: str
    sheet: str
    header: Dict[str, Field] = field(default_factory=dict)
    rows: List[DetailRow] = field(default_factory=list)
    total: Optional[Field] = None
    issues: List[Issue] = field(default_factory=list)
    table_found: bool = False
    available_columns: List[str] = field(default_factory=list)
    row_snapshots: List[RowSnapshot] = field(default_factory=list)
    #: 额外费用结构候选（折扣/运费/税费等），供 R05 判断金额结构是否受支持
    extra_cost_rows: List[Dict[str, Any]] = field(default_factory=list)

    def header_field(self, name: str) -> Field:
        return self.header.get(name) or Field(name=name)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "role": self.role,
            "role_label": self.role_label,
            "file_name": self.file_name,
            "sha256": self.sha256,
            "template_id": self.template_id,
            "sheet": self.sheet,
            "header": {k: v.to_dict() for k, v in self.header.items()},
            "rows": [r.to_dict() for r in self.rows],
            "total": self.total.to_dict() if self.total else None,
            "issues": [i.to_dict() for i in self.issues],
            "table_found": self.table_found,
            "available_columns": list(self.available_columns),
            "row_snapshots": [s.to_dict() for s in self.row_snapshots],
            "extra_cost_rows": list(self.extra_cost_rows),
        }


# ---------------------------------------------------------------- 检查结果


@dataclass
class Finding:
    """一条检查结果。任务书 5. 要求每条都含规则编号/项目/状态/两边值/差异说明/文件名/工作表/单元格。"""

    rule_id: str
    item: str
    status: str
    expected: str = ""
    actual: str = ""
    evidence: str = ""
    detail: str = ""
    sources: List[CellRef] = field(default_factory=list)
    corrections: List[str] = field(default_factory=list)

    def source_text(self) -> str:
        return "；".join(s.text() for s in self.sources) if self.sources else "—"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "rule_id": self.rule_id,
            "item": self.item,
            "status": self.status,
            "expected": self.expected,
            "actual": self.actual,
            "evidence": self.evidence,
            "detail": self.detail,
            "sources": [s.to_dict() for s in self.sources],
            "corrections": list(self.corrections),
        }


@dataclass
class Correction:
    """用户对本次会话提取值的修正记录（不改动原文件）。"""

    target: str  # 例如 "invoice:row.9.qty"
    old_text: str
    new_text: str
    source: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "target": self.target,
            "old_text": self.old_text,
            "new_text": self.new_text,
            "source": self.source,
        }


@dataclass
class CheckResult:
    findings: List[Finding] = field(default_factory=list)
    overall: str = ""
    overall_note: str = ""
    coverage: List[str] = field(default_factory=list)
    unchecked: List[str] = field(default_factory=list)
    corrections: List[Correction] = field(default_factory=list)

    def count(self, status: str) -> int:
        return sum(1 for f in self.findings if f.status == status)

    def summary(self) -> Dict[str, int]:
        return {
            STATUS_DIFF: self.count(STATUS_DIFF),
            STATUS_TODO: self.count(STATUS_TODO),
            STATUS_SKIP: self.count(STATUS_SKIP),
            STATUS_OK: self.count(STATUS_OK),
        }

    def to_dict(self) -> Dict[str, Any]:
        return {
            "findings": [f.to_dict() for f in self.findings],
            "overall": self.overall,
            "overall_note": self.overall_note,
            "coverage": list(self.coverage),
            "unchecked": list(self.unchecked),
            "corrections": [c.to_dict() for c in self.corrections],
            "summary": self.summary(),
        }
