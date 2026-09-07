"""一致性检查规则 R01 - R06。

规则清单
  R01 文件关联与检查范围      -> check_reference_consistency
  R02 产品对应关系            -> build_match_groups / check_quantity
  R03 数量差异与缺失产品      -> check_quantity
  R04 发票行金额              -> check_line_amounts
  R05 发票总额                -> check_invoice_total
  R06 输入完整性              -> check_input_integrity

状态口径（任务书 5.）
  未发现差异：输入及规则具备检查条件，检查完成且无差异。
  发现差异  ：有可复现的比较或计算依据。
  待确认    ：识别、匹配或业务含义不明确。
  未检查    ：超出支持范围或缺少检查条件。

关键约束
  金额一律用 Decimal，不做二进制浮点比较（R04）。
  公式单元格的缓存值不视为可靠数值，凡依赖它的检查一律降级为「未检查」（R06）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Dict, List, Optional, Tuple

from .models import (
    FIELD_MISSING,
    FIELD_NEGATIVE,
    FIELD_NON_NUMERIC,
    FIELD_OK,
    FIELD_UNVERIFIED,
    STATUS_DIFF,
    STATUS_OK,
    STATUS_SKIP,
    STATUS_TODO,
    CellRef,
    CheckResult,
    Correction,
    DetailRow,
    DocumentExtraction,
    Field,
    Finding,
    worst_status,
)
from .money import (
    format_decimal,
    parse_currency,
    quantize_amount,
    to_decimal,
    within_tolerance,
)

#: v0.1 不支持任何单位换算关系。件 ≠ 箱，也不猜测比例（R03）。
UNIT_EQUIVALENCES: Dict[Tuple[str, str], Decimal] = {}


# ---------------------------------------------------------------- 分组结构


@dataclass
class Group:
    """按（型号, 规格）汇总后的一个产品条目。"""

    key: Tuple[str, str]
    model_texts: List[str] = field(default_factory=list)
    spec_texts: List[str] = field(default_factory=list)
    unit_keys: set = field(default_factory=set)
    unit_texts: List[str] = field(default_factory=list)
    qty: Optional[Decimal] = None
    qty_parts: List[str] = field(default_factory=list)
    cells: List[CellRef] = field(default_factory=list)
    row_ids: List[int] = field(default_factory=list)
    problems: List[str] = field(default_factory=list)
    degraded: bool = False

    @property
    def model_display(self) -> str:
        return self.model_texts[0] if self.model_texts else "（缺失）"

    @property
    def spec_display(self) -> str:
        return self.spec_texts[0] if self.spec_texts else "—"

    @property
    def unit_display(self) -> str:
        uniq = sorted({u for u in self.unit_texts if u})
        return "、".join(uniq) if uniq else "（缺失）"


def _model_missing(f: Field) -> bool:
    return f.state == FIELD_MISSING or not f.text.strip()


def _build_groups(doc: DocumentExtraction, use_spec: bool) -> Tuple[Dict[Tuple[str, str], Group], List[DetailRow]]:
    """把明细行按 (型号, 规格) 汇总。型号缺失的行单独返回，不参与匹配。"""
    groups: Dict[Tuple[str, str], Group] = {}
    skipped: List[DetailRow] = []

    for row in doc.rows:
        model = row.get("model")
        if _model_missing(model):
            skipped.append(row)
            continue
        spec = row.get("spec")
        spec_key = spec.key if (use_spec and spec.text.strip()) else ""
        key = (model.key, spec_key)

        g = groups.get(key)
        if g is None:
            g = Group(key=key)
            groups[key] = g
        if model.text not in g.model_texts:
            g.model_texts.append(model.text)
        if spec.text and spec.text not in g.spec_texts:
            g.spec_texts.append(spec.text)

        unit = row.get("unit")
        if unit.state == FIELD_OK and unit.key:
            g.unit_keys.add(unit.key)
            if unit.text not in g.unit_texts:
                g.unit_texts.append(unit.text)
        elif unit.state != FIELD_MISSING and unit.text:
            if unit.text not in g.unit_texts:
                g.unit_texts.append(unit.text)

        qty = row.get("qty")
        g.row_ids.append(row.row_index)
        if qty.cell:
            g.cells.append(qty.cell)
        if model.cell and model.cell not in g.cells:
            g.cells.append(model.cell)
        for f in (spec, unit):
            if f.cell and f.cell not in g.cells:
                g.cells.append(f.cell)

        if qty.state == FIELD_MISSING:
            g.degraded = True
            g.problems.append(f"第 {row.row_index} 行数量缺失")
        elif qty.state == FIELD_NON_NUMERIC:
            g.degraded = True
            g.problems.append(f"第 {row.row_index} 行数量不是数字：{qty.text!r}")
        elif qty.state == FIELD_NEGATIVE:
            g.degraded = True
            g.problems.append(f"第 {row.row_index} 行数量为负数：{qty.text}")
        elif qty.state == FIELD_UNVERIFIED:
            g.degraded = True
            g.problems.append(
                f"第 {row.row_index} 行数量为公式且缓存未验证：{qty.note}"
            )
        else:
            val = qty.number
            g.qty = (g.qty or Decimal(0)) + (val if val is not None else Decimal(0))
            g.qty_parts.append(f"{qty.text}（{qty.cell.coord if qty.cell else '?'}）")

        if unit.state == FIELD_MISSING:
            g.degraded = True
            g.problems.append(f"第 {row.row_index} 行单位缺失")
        if spec.state == FIELD_UNVERIFIED:
            g.degraded = True
            g.problems.append(f"第 {row.row_index} 行规格为公式且缓存未验证")
        if unit.state == FIELD_UNVERIFIED:
            g.degraded = True
            g.problems.append(f"第 {row.row_index} 行单位为公式且缓存未验证：{unit.note}")
        if model.state == FIELD_UNVERIFIED:
            g.degraded = True
            g.problems.append(f"第 {row.row_index} 行型号为公式且缓存未验证：{model.note}")

    return groups, skipped


# ---------------------------------------------------------------- R01


def check_reference_consistency(
    invoice: DocumentExtraction, packing: DocumentExtraction
) -> List[Finding]:
    """R01：检查双方同类引用字段是否一致。

    不要求两种单据的自身编号相同，也不比较签发日期、公司名称。
    """
    findings: List[Finding] = []
    pairs = [
        ("order_ref", "order_ref", "订单引用号"),
        ("doc_no", "invoice_ref", "发票号（装箱单引用）"),
    ]
    compared = 0
    for inv_name, pac_name, label in pairs:
        a = invoice.header_field(inv_name)
        b = packing.header_field(pac_name)
        if a.state == FIELD_MISSING or b.state == FIELD_MISSING:
            continue
        compared += 1
        if a.state == FIELD_UNVERIFIED or b.state == FIELD_UNVERIFIED:
            findings.append(
                Finding(
                    rule_id="R01",
                    item=f"{label}一致性",
                    status=STATUS_SKIP,
                    expected=f"发票“{a.text}”",
                    actual=f"装箱单“{b.text}”",
                    evidence="引用字段含公式，缓存值未验证，不能作为关联依据。",
                    detail=f"{label}为公式且缓存未验证（R06），本项未检查。",
                    sources=[c for c in (a.cell, b.cell) if c],
                )
            )
            continue
        if a.key == b.key:
            findings.append(
                Finding(
                    rule_id="R01",
                    item=f"{label}一致性",
                    status=STATUS_OK,
                    expected=f"发票“{a.text}”",
                    actual=f"装箱单“{b.text}”",
                    evidence="两边规范化后一致。",
                    detail=f"{label}一致，可判定两份单据指向同一批出货（仅此项）。",
                    sources=[c for c in (a.cell, b.cell) if c],
                )
            )
        else:
            findings.append(
                Finding(
                    rule_id="R01",
                    item=f"{label}一致性",
                    status=STATUS_DIFF,
                    expected=f"发票“{a.text}”",
                    actual=f"装箱单“{b.text}”",
                    evidence="两边规范化后不一致。",
                    detail=f"{label}不一致，两份单据可能不属于同一批出货，需人工确认后再核对。",
                    sources=[c for c in (a.cell, b.cell) if c],
                )
            )

    if compared == 0:
        findings.append(
            Finding(
                rule_id="R01",
                item="单据关联字段",
                status=STATUS_TODO,
                expected="至少一对可比较的引用字段（订单号 或 发票号）",
                actual="两边均未同时提供",
                evidence=(
                    f"发票：订单引用号={invoice.header_field('order_ref').text or '（空）'}，"
                    f"发票号={invoice.header_field('doc_no').text or '（空）'}；"
                    f"装箱单：订单引用号={packing.header_field('order_ref').text or '（空）'}，"
                    f"发票引用={packing.header_field('invoice_ref').text or '（空）'}"
                ),
                detail="缺少可比较的引用字段，无法确认两份单据属于同一批出货，需人工确认。",
                sources=[c for c in (invoice.header_field('doc_no').cell, packing.header_field('invoice_ref').cell) if c],
            )
        )
    return findings


# ---------------------------------------------------------------- R02 / R03


def check_quantity(
    invoice: DocumentExtraction, packing: DocumentExtraction
) -> Tuple[List[Finding], List[str], List[str]]:
    """R02 + R03：建立产品对应关系并比较汇总数量。

    返回 (findings, coverage, unchecked)。
    """
    findings: List[Finding] = []
    coverage: List[str] = []
    unchecked: List[str] = []

    def _has_spec(doc: DocumentExtraction) -> bool:
        return any(
            r.get("spec").state != FIELD_MISSING and r.get("spec").text.strip()
            for r in doc.rows
        )

    inv_has_spec = _has_spec(invoice)
    pac_has_spec = _has_spec(packing)
    use_spec = inv_has_spec and pac_has_spec

    # 只有「一侧有规格、另一侧没有」才算歧义；两侧都没有规格值时不算问题
    specs_by_model = {}
    for doc in (invoice, packing):
        for row in doc.rows:
            specs_by_model.setdefault(row.get("model").key, []).append(row.get("spec"))
    ambiguous_specs = any(
        any(f.state == FIELD_UNVERIFIED for f in fields)
        or (any(f.text.strip() for f in fields) and any(not f.text.strip() for f in fields))
        for fields in specs_by_model.values()
    )
    if inv_has_spec != pac_has_spec or ambiguous_specs:
        findings.append(
            Finding(
                rule_id="R02",
                item="匹配键：规格字段",
                status=STATUS_TODO,
                expected="两边明细都提供规格列",
                actual=(
                    f"发票有规格={'是' if inv_has_spec else '否'}，"
                    f"装箱单有规格={'是' if pac_has_spec else '否'}"
                ),
                evidence="部分规格缺失或含未验证公式，无法区分同型号不同规格的产品。",
                detail="本次数量匹配未执行；请补齐或确认规格后重新核对，不将不同规格合并比较。",
                sources=[],
            )
        )
        unchecked.append("数量比较（规格不完整或未验证，匹配关系待确认）")
        return findings, coverage, unchecked

    inv_groups, inv_skipped = _build_groups(invoice, use_spec)
    pac_groups, pac_skipped = _build_groups(packing, use_spec)

    for row in inv_skipped:
        findings.append(
            Finding(
                rule_id="R02",
                item=f"发票第 {row.row_index} 行型号",
                status=STATUS_TODO,
                expected="型号非空",
                actual="（空）",
                evidence="型号缺失，无法建立与装箱单的对应关系。",
                detail="该行未参与数量比较，需人工补充型号后重新检查。",
                sources=[c for c in (row.get("model").cell,) if c],
            )
        )
    for row in pac_skipped:
        findings.append(
            Finding(
                rule_id="R02",
                item=f"装箱单第 {row.row_index} 行型号",
                status=STATUS_TODO,
                expected="型号非空",
                actual="（空）",
                evidence="型号缺失，无法建立与发票的对应关系。",
                detail="该行未参与数量比较，需人工补充型号后重新检查。",
                sources=[c for c in (row.get("model").cell,) if c],
            )
        )

    all_keys = list(inv_groups.keys()) + [k for k in pac_groups if k not in inv_groups]
    # 稳定排序：先按型号再按规格，保证报告顺序可复现
    all_keys.sort(key=lambda k: (k[0], k[1]))

    for key in all_keys:
        gi = inv_groups.get(key)
        gp = pac_groups.get(key)
        label = f"{gi.model_display if gi else gp.model_display}" + (
            f" / {gi.spec_display if gi else gp.spec_display}" if use_spec and (gi.spec_display if gi else gp.spec_display) != "—" else ""
        )

        if (gi and gi.degraded) or (gp and gp.degraded):
            groups = [g for g in (gi, gp) if g]
            findings.append(Finding(
                rule_id="R03", item=f"数量比较：{label}", status=STATUS_SKIP,
                expected="型号、规格、单位和数量均可确认",
                actual="匹配或数量输入不可用",
                evidence="；".join(p for g in groups for p in g.problems),
                detail="相关数据缺失或含未验证公式，未执行数量与产品缺失判断。",
                sources=[c for g in groups for c in g.cells],
            ))
            unchecked.append(f"数量比较：{label}（数据不完整或未验证）")
            continue

        if gi is not None and gp is None:
            findings.append(
                Finding(
                    rule_id="R03",
                    item=f"仅发票出现的产品：{label}",
                    status=STATUS_DIFF,
                    expected=f"装箱单应有 {format_decimal(gi.qty)}（发票汇总）",
                    actual="装箱单无此产品",
                    evidence="；".join(gi.qty_parts) or "发票汇总数量不可用",
                    detail="该产品只在商业发票中出现，装箱单缺失。",
                    sources=gi.cells,
                )
            )
            continue
        if gp is not None and gi is None:
            findings.append(
                Finding(
                    rule_id="R03",
                    item=f"仅装箱单出现的产品：{label}",
                    status=STATUS_DIFF,
                    expected="发票无此产品",
                    actual=f"装箱单汇总 {format_decimal(gp.qty)}",
                    evidence="；".join(gp.qty_parts) or "装箱单汇总数量不可用",
                    detail="该产品只在装箱单中出现，发票缺失。",
                    sources=gp.cells,
                )
            )
            continue

        # 两边都有：先核对单位
        if not gi.unit_keys or not gp.unit_keys:
            findings.append(
                Finding(
                    rule_id="R03",
                    item=f"单位：{label}",
                    status=STATUS_TODO,
                    expected=f"发票单位 {gi.unit_display}",
                    actual=f"装箱单单位 {gp.unit_display}",
                    evidence="一侧或两侧单位缺失/为公式未验证，无法确认单位一致。",
                    detail="单位无法确认一致，数量未比较。本工具不把「箱」当成「件」，也不猜测换算比例。",
                    sources=gi.cells + gp.cells,
                )
            )
            unchecked.append(f"数量比较：{label}（单位缺失或未验证）")
            continue

        if gi.unit_keys != gp.unit_keys or len(gi.unit_keys) != 1 or len(gp.unit_keys) != 1:
            findings.append(
                Finding(
                    rule_id="R03",
                    item=f"单位不一致：{label}",
                    status=STATUS_TODO,
                    expected=f"发票单位 {gi.unit_display}",
                    actual=f"装箱单单位 {gp.unit_display}",
                    evidence=(
                        f"发票汇总 {format_decimal(gi.qty)} {gi.unit_display}；"
                        f"装箱单汇总 {format_decimal(gp.qty)} {gp.unit_display}"
                    ),
                    detail=(
                        "两边单位不一致，数量不直接比较（R03：不把「箱」当成「件」，不自动猜测换算比例）。"
                        "v0.1 不支持任何单位换算关系，需人工确认后自行核对。"
                    ),
                    sources=gi.cells + gp.cells,
                )
            )
            unchecked.append(f"数量比较：{label}（单位不一致，未比较）")
            continue

        # 单位一致，比较汇总数量
        problems = gi.problems + gp.problems
        if gi.degraded or gp.degraded or gi.qty is None or gp.qty is None:
            findings.append(
                Finding(
                    rule_id="R03",
                    item=f"数量比较：{label}",
                    status=STATUS_SKIP,
                    expected=f"发票汇总 {format_decimal(gi.qty)} {gi.unit_display}",
                    actual=f"装箱单汇总 {format_decimal(gp.qty)} {gp.unit_display}",
                    evidence="；".join(problems) or "汇总数量不可用",
                    detail="数量数据不完整或含未验证公式缓存，未执行比较（R06）。",
                    sources=gi.cells + gp.cells,
                )
            )
            unchecked.append(f"数量比较：{label}（数据不完整）")
            continue

        if gi.qty == gp.qty:
            findings.append(
                Finding(
                    rule_id="R03",
                    item=f"数量比较：{label}",
                    status=STATUS_OK,
                    expected=f"发票 {format_decimal(gi.qty)} {gi.unit_display}",
                    actual=f"装箱单 {format_decimal(gp.qty)} {gp.unit_display}",
                    evidence=(
                        f"发票汇总：{' + '.join(gi.qty_parts)} = {format_decimal(gi.qty)}；"
                        f"装箱单汇总：{' + '.join(gp.qty_parts)} = {format_decimal(gp.qty)}"
                    ),
                    detail=f"按（型号{', 规格' if use_spec else ''}）汇总后数量一致。",
                    sources=gi.cells + gp.cells,
                )
            )
        else:
            diff = gp.qty - gi.qty
            findings.append(
                Finding(
                    rule_id="R03",
                    item=f"数量差异：{label}",
                    status=STATUS_DIFF,
                    expected=f"发票 {format_decimal(gi.qty)} {gi.unit_display}",
                    actual=f"装箱单 {format_decimal(gp.qty)} {gp.unit_display}",
                    evidence=(
                        f"发票汇总：{' + '.join(gi.qty_parts)} = {format_decimal(gi.qty)}；"
                        f"装箱单汇总：{' + '.join(gp.qty_parts)} = {format_decimal(gp.qty)}；"
                        f"差额（装箱单 - 发票）= {format_decimal(diff)} {gi.unit_display}"
                    ),
                    detail=f"汇总数量不一致，差额 {format_decimal(diff)} {gi.unit_display}（装箱单相对发票）。",
                    sources=gi.cells + gp.cells,
                )
            )
        coverage.append(f"数量比较：{label}（单位 {gi.unit_display}）")

    return findings, coverage, unchecked


# ---------------------------------------------------------------- R04


def check_line_amounts(
    invoice: DocumentExtraction, money: Dict[str, Any], currency: Dict[str, Any]
) -> Tuple[List[Finding], List[str], List[str]]:
    """R04：数量 x 单价 = 行金额。

    money    : 模板里的舍入与容差配置
    currency : {"code", "digits", "ok", "message"}
    """
    findings: List[Finding] = []
    coverage: List[str] = []
    unchecked: List[str] = []

    digits = currency.get("digits")
    code = currency.get("code") or "（未知）"
    mode = money["rounding_mode"]
    tolerance = to_decimal(money["tolerance_absolute"])
    if tolerance is None or tolerance < 0:
        raise ValueError("金额容差必须是非负数")

    for row in invoice.rows:
        qty = row.get("qty")
        price = row.get("unit_price")
        amount = row.get("amount")
        label = row.get("model").text or f"第 {row.row_index} 行"
        item = f"发票第 {row.row_index} 行行金额（{label}）"
        cells = [c for c in (qty.cell, price.cell, amount.cell) if c]

        blocking: List[str] = []
        for name, f in (("数量", qty), ("单价", price), ("行金额", amount)):
            if f.state == FIELD_MISSING:
                blocking.append(f"{name}缺失")
            elif f.state == FIELD_NON_NUMERIC:
                blocking.append(f"{name}不是数字（{f.text!r}）")
            elif f.state == FIELD_NEGATIVE:
                blocking.append(f"{name}为负数（{f.text}）")
            elif f.state == FIELD_UNVERIFIED:
                blocking.append(f"{name}为公式且缓存未验证")

        if blocking:
            status = STATUS_SKIP if any("公式" in b for b in blocking) else STATUS_TODO
            findings.append(
                Finding(
                    rule_id="R04",
                    item=item,
                    status=status,
                    expected="数量 × 单价（按配置舍入后）",
                    actual=f"数量={qty.text or '（空）'}，单价={price.text or '（空）'}，行金额={amount.text or '（空）'}",
                    evidence="；".join(blocking),
                    detail="缺少可复现的计算条件，未执行行金额检查（R06）。" if status == STATUS_SKIP
                           else "输入值不可用，需人工确认后重新检查（R06）。",
                    sources=cells,
                )
            )
            unchecked.append(item)
            continue

        if digits is None:
            findings.append(
                Finding(
                    rule_id="R04",
                    item=item,
                    status=STATUS_SKIP,
                    expected="需先确定币种及小数位数规则",
                    actual=f"币种={code}",
                    evidence=currency.get("message", ""),
                    detail="币种规则未知，本工具不假定所有币种都是两位小数，未执行行金额检查。",
                    sources=cells,
                )
            )
            unchecked.append(item)
            continue

        expected = quantize_amount(qty.number * price.number, digits, mode)
        diff = amount.number - expected
        expr = (
            f"{format_decimal(qty.number)} × {format_decimal(price.number, max(digits, -price.number.as_tuple().exponent))} "
            f"= {format_decimal(qty.number * price.number, digits + 4)}，"
            f"按 {mode} 保留 {digits} 位 -> {format_decimal(expected, digits)} {code}；"
            f"单据行金额 {format_decimal(amount.number, digits)}；"
            f"差 {format_decimal(diff, digits)}（容差 {format_decimal(tolerance, digits)}）"
        )

        if within_tolerance(expected, amount.number, tolerance):
            findings.append(
                Finding(
                    rule_id="R04",
                    item=item,
                    status=STATUS_OK,
                    expected=f"{format_decimal(expected, digits)} {code}",
                    actual=f"{format_decimal(amount.number, digits)} {code}",
                    evidence=expr,
                    detail="行金额与数量×单价一致（在配置容差内）。",
                    sources=cells,
                )
            )
        else:
            findings.append(
                Finding(
                    rule_id="R04",
                    item=item,
                    status=STATUS_DIFF,
                    expected=f"{format_decimal(expected, digits)} {code}",
                    actual=f"{format_decimal(amount.number, digits)} {code}",
                    evidence=expr,
                    detail=f"行金额与计算结果相差 {format_decimal(diff, digits)} {code}，超出容差。",
                    sources=cells,
                )
            )
        coverage.append(item)

    return findings, coverage, unchecked


# ---------------------------------------------------------------- R05


def check_invoice_total(
    invoice: DocumentExtraction, money: Dict[str, Any], currency: Dict[str, Any]
) -> Tuple[List[Finding], List[str], List[str]]:
    """R05：明细金额合计 vs 单据总额；并识别额外费用结构。"""
    findings: List[Finding] = []
    coverage: List[str] = []
    unchecked: List[str] = []

    digits = currency.get("digits")
    code = currency.get("code") or "（未知）"
    tolerance = to_decimal(money["tolerance_absolute"])
    if tolerance is None or tolerance < 0:
        raise ValueError("金额容差必须是非负数")

    extras = invoice.extra_cost_rows or []
    if extras:
        ev = "；".join(
            f"“{e['label']}”（{e['label_cell']['sheet']}!{e['label_cell']['coord']}）"
            + "".join(
                f" 值 {n['text']} @ {n['cell']['sheet']}!{n['cell']['coord']}"
                for n in e["numbers"][:3]
            )
            for e in extras
        )
        findings.append(
            Finding(
                rule_id="R05",
                item="金额结构：额外费用/折扣/税费",
                status=STATUS_TODO,
                expected="模板要求：无折扣、无税费、无附加费用",
                actual=f"发现 {len(extras)} 处疑似额外费用行",
                evidence=ev,
                detail="不支持该金额结构，需人工核对。本工具不会忽略这些项目后宣称总额通过。",
                sources=[
                    CellRef(
                        file_name=invoice.file_name,
                        sheet=e["label_cell"]["sheet"],
                        coord=e["label_cell"]["coord"],
                    )
                    for e in extras
                ],
            )
        )

    total = invoice.total
    amounts: List[Field] = []
    bad: List[str] = []
    for row in invoice.rows:
        a = row.get("amount")
        if a.state == FIELD_MISSING:
            bad.append(f"第 {row.row_index} 行金额缺失")
        elif a.state == FIELD_NON_NUMERIC:
            bad.append(f"第 {row.row_index} 行金额不是数字（{a.text!r}）")
        elif a.state == FIELD_UNVERIFIED:
            bad.append(f"第 {row.row_index} 行金额为公式且缓存未验证")
        elif a.state == FIELD_NEGATIVE:
            bad.append(f"第 {row.row_index} 行金额为负数（{a.text}）")
        else:
            amounts.append(a)

    if total is None or total.state == FIELD_MISSING:
        findings.append(
            Finding(
                rule_id="R05",
                item="发票总额",
                status=STATUS_TODO,
                expected="应存在总额行",
                actual="未找到总额行或总额为空",
                evidence="按模板配置的合计标签未能定位到非空总额单元格。",
                detail="缺少单据总额，无法完成总额检查，需人工确认。",
                sources=[],
            )
        )
        unchecked.append("发票总额合计检查")
        return findings, coverage, unchecked

    if total.state in (FIELD_NON_NUMERIC, FIELD_NEGATIVE) or not invoice.rows:
        findings.append(Finding(
            rule_id="R05", item="发票总额", status=STATUS_TODO,
            expected="有效的非负总额与非空明细", actual=total.text or "（空）",
            evidence="总额无效或无明细，不能执行合计计算。",
            detail="请修正总额或确认明细后重新核对。",
            sources=[total.cell] if total.cell else [],
        ))
        unchecked.append("发票总额合计检查（输入无效）")
        return findings, coverage, unchecked

    if extras:
        findings.append(Finding(
            rule_id="R05", item="发票总额", status=STATUS_TODO,
            expected="无附加费用的单一金额结构", actual=total.text,
            evidence="存在额外费用，不能仅凭明细相加相等判定总额通过。",
            detail="不支持该金额结构，需人工核对。",
            sources=[total.cell] if total.cell else [],
        ))
        unchecked.append("发票总额合计检查（额外费用结构）")
        return findings, coverage, unchecked

    if bad or digits is None:
        reason = "；".join(bad) if bad else currency.get("message", "币种规则未知")
        findings.append(
            Finding(
                rule_id="R05",
                item="发票总额",
                status=STATUS_SKIP,
                expected="明细金额合计 = 单据总额",
                actual=f"单据总额 {total.text or '（空）'}",
                evidence=reason,
                detail="明细金额不完整或币种规则未知，未执行总额比较（R06）。",
                sources=[c for c in (total.cell,) if c],
            )
        )
        unchecked.append("发票总额合计检查")
        return findings, coverage, unchecked

    if total.state == FIELD_UNVERIFIED:
        findings.append(
            Finding(
                rule_id="R05",
                item="发票总额",
                status=STATUS_SKIP,
                expected="明细金额合计 = 单据总额",
                actual=f"总额单元格为公式且缓存未验证（缓存值 {total.text}）",
                evidence=total.note,
                detail="总额为公式缓存值，未重新计算，不标记为通过（R06）。",
                sources=[c for c in (total.cell,) if c],
            )
        )
        unchecked.append("发票总额合计检查")
        return findings, coverage, unchecked

    line_sum = sum((a.number for a in amounts), Decimal(0))
    diff = total.number - line_sum
    expr = (
        " + ".join(
            f"{format_decimal(a.number, digits)}（{a.cell.coord if a.cell else '?'}）" for a in amounts
        )
        + f" = {format_decimal(line_sum, digits)}；"
        f"单据总额 {format_decimal(total.number, digits)}"
        f"（{total.cell.coord if total.cell else '?'}）；"
        f"差 {format_decimal(diff, digits)}（容差 {format_decimal(tolerance, digits)}）"
    )
    cells = [a.cell for a in amounts if a.cell] + ([total.cell] if total.cell else [])

    if within_tolerance(line_sum, total.number, tolerance):
        findings.append(
            Finding(
                rule_id="R05",
                item="发票总额",
                status=STATUS_OK,
                expected=f"{format_decimal(line_sum, digits)} {code}",
                actual=f"{format_decimal(total.number, digits)} {code}",
                evidence=expr,
                detail="明细金额合计与单据总额一致（在配置容差内）。",
                sources=cells,
            )
        )
    else:
        if extras:
            findings.append(
                Finding(
                    rule_id="R05",
                    item="发票总额",
                    status=STATUS_TODO,
                    expected=f"{format_decimal(line_sum, digits)} {code}",
                    actual=f"{format_decimal(total.number, digits)} {code}",
                    evidence=expr + "；同时发现额外费用结构",
                    detail=(
                        f"明细合计与总额相差 {format_decimal(diff, digits)} {code}，"
                        "且存在额外费用/折扣/税费行，可能属于本工具不支持的金额结构，需人工核对。"
                    ),
                    sources=cells,
                )
            )
            unchecked.append("发票总额合计检查（存在额外费用结构，结论待人工确认）")
        else:
            findings.append(
                Finding(
                    rule_id="R05",
                    item="发票总额",
                    status=STATUS_DIFF,
                    expected=f"{format_decimal(line_sum, digits)} {code}",
                    actual=f"{format_decimal(total.number, digits)} {code}",
                    evidence=expr,
                    detail=(
                        f"明细金额合计与单据总额相差 {format_decimal(diff, digits)} {code}，"
                        "且未发现可解释的额外费用行。"
                    ),
                    sources=cells,
                )
            )
        coverage.append("发票总额合计检查")
        return findings, coverage, unchecked

    coverage.append("发票总额合计检查")
    return findings, coverage, unchecked


# ---------------------------------------------------------------- R06


def check_input_integrity(
    invoice: DocumentExtraction, packing: DocumentExtraction, template: Dict[str, Any]
) -> Tuple[List[Finding], List[str]]:
    """R06：输入完整性。把提取阶段的问题转成带规则编号的结论。"""
    findings: List[Finding] = []
    unchecked: List[str] = []

    for doc in (invoice, packing):
        prefix = doc.role_label
        for issue in doc.issues:
            findings.append(
                Finding(
                    rule_id="R06",
                    item=f"{prefix}：{issue.code}",
                    status=issue.severity,
                    expected="输入可被完整读取",
                    actual=issue.message,
                    evidence=issue.cell.text() if issue.cell else doc.file_name,
                    detail=(
                        "该项未纳入本次核对范围，需人工处理。"
                        if issue.severity == STATUS_SKIP
                        else "需人工确认后重新检查。"
                    ),
                    sources=[issue.cell] if issue.cell else [],
                )
            )
            unchecked.append(f"{prefix}：{issue.message}")

        if not doc.rows:
            continue

        for row in doc.rows:
            for name in ("model", "unit", "qty"):
                f = row.get(name)
                if f.state in (FIELD_MISSING, FIELD_NON_NUMERIC, FIELD_NEGATIVE):
                    findings.append(
                        Finding(
                            rule_id="R06",
                            item=f"{prefix} 第 {row.row_index} 行{name}",
                            status=STATUS_TODO,
                            expected=f"{name} 应为有效值",
                            actual=f.text or "（空）",
                            evidence=f.state,
                            detail="必需字段缺失或不可用，该行相关检查未执行。",
                            sources=[f.cell] if f.cell else [],
                        )
                    )
            for name, f in row.fields.items():
                if f.state == FIELD_UNVERIFIED:
                    findings.append(
                        Finding(
                            rule_id="R06",
                            item=f"{prefix} 第 {row.row_index} 行{name}",
                            status=STATUS_SKIP,
                            expected="可直接读取的常量值",
                            actual=f"公式 {f.formula}（缓存值 {f.text or '无'}）",
                            evidence=f.note,
                            detail="公式缓存未重新计算，未作为可靠数值使用，相关检查降级。",
                            sources=[f.cell] if f.cell else [],
                        )
                    )

    return findings, unchecked


# ---------------------------------------------------------------- 主入口


def resolve_currency(invoice: DocumentExtraction, money: Dict[str, Any]) -> Dict[str, Any]:
    """确定发票币种与小数位数规则。未知币种一律降级，不假定两位小数。"""
    raw_field = invoice.header_field("currency")
    text = raw_field.text or ""
    if raw_field.state == FIELD_UNVERIFIED:
        return {
            "code": None,
            "digits": None,
            "ok": False,
            "message": f"币种字段为公式且缓存未验证：{raw_field.formula}",
            "raw": text,
            "cell": raw_field.cell,
        }
    if not text:
        return {
            "code": None,
            "digits": None,
            "ok": False,
            "message": "发票币种字段为空，无法确定金额小数位数规则",
            "raw": text,
            "cell": raw_field.cell,
        }
    code, digits, msg = parse_currency(text)
    if code is None or digits is None:
        return {
            "code": None,
            "digits": None,
            "ok": False,
            "message": msg,
            "raw": text,
            "cell": raw_field.cell,
        }
    return {"code": code, "digits": digits, "ok": True, "message": msg, "raw": text, "cell": raw_field.cell}


def run_checks(
    invoice: DocumentExtraction,
    packing: DocumentExtraction,
    template: Dict[str, Any],
    corrections: Optional[List[Correction]] = None,
) -> CheckResult:
    """执行全部检查，返回 CheckResult。"""
    from .templates import money_config

    money = money_config(template)
    currency = resolve_currency(invoice, money)

    result = CheckResult(corrections=list(corrections or []))
    coverage: List[str] = []
    unchecked: List[str] = []

    result.findings += check_reference_consistency(invoice, packing)

    if not currency["ok"]:
        result.findings.append(
            Finding(
                rule_id="R04",
                item="币种与金额规则",
                status=STATUS_TODO,
                expected="可识别的币种及其小数位数规则",
                actual=f"币种字段：{currency['raw'] or '（空）'}",
                evidence=currency["message"],
                detail=(
                    "币种未识别或规则未知，行金额与总额检查未执行。"
                    "本工具不假定所有币种都是两位小数，也不混合币种计算。"
                ),
                sources=[currency["cell"]] if currency.get("cell") else [],
            )
        )
        unchecked.append("行金额检查（币种未知）")
        unchecked.append("发票总额检查（币种未知）")

    if invoice.table_found and packing.table_found:
        f, cov, unc = check_quantity(invoice, packing)
        result.findings += f
        coverage += cov
        unchecked += unc
    else:
        result.findings.append(
            Finding(
                rule_id="R03",
                item="数量核对",
                status=STATUS_SKIP,
                expected="两边明细表均可识别",
                actual=(
                    f"发票明细表可识别={'是' if invoice.table_found else '否'}，"
                    f"装箱单明细表可识别={'是' if packing.table_found else '否'}"
                ),
                evidence="明细表未能按所选模板定位，缺少检查条件。",
                detail="未执行数量核对。请确认模板选择正确，或改用受支持的模板。",
                sources=[],
            )
        )
        unchecked.append("数量核对（明细表未识别）")

    if invoice.table_found:
        if currency["ok"]:
            f, cov, unc = check_line_amounts(invoice, money, currency)
            result.findings += f
            coverage += cov
            unchecked += unc
        f, cov, unc = check_invoice_total(invoice, money, currency)
        result.findings += f
        coverage += cov
        unchecked += unc
    else:
        unchecked.append("行金额检查（发票明细表未识别）")
        unchecked.append("发票总额检查（发票明细表未识别）")

    f, unc = check_input_integrity(invoice, packing, template)
    result.findings += f
    unchecked += unc

    result.coverage = coverage
    result.unchecked = unchecked
    result.overall, result.overall_note = _overall(result)
    return result


def _overall(result: CheckResult) -> Tuple[str, str]:
    """总体结论用有限表述（任务书 5.）。"""
    n_diff = result.count(STATUS_DIFF)
    n_todo = result.count(STATUS_TODO)
    n_skip = result.count(STATUS_SKIP)

    if n_diff:
        head = f"在已检查范围内发现 {n_diff} 项差异"
    else:
        head = "在已检查范围内未发现差异"

    notes: List[str] = []
    if n_todo:
        notes.append(f"{n_todo} 项待确认")
    if n_skip:
        notes.append(f"{n_skip} 项未检查")
    tail = ("，另有 " + "、".join(notes)) if notes else ""

    note = (
        "本结论只覆盖本次实际执行的检查项，不覆盖“未检查/待确认”的部分。"
        "重要金额、币种及出货数据仍需用户人工复核。"
    )
    if result.corrections:
        note += f" 本报告基于用户修正后的数据（{len(result.corrections)} 处修正），非原始提取值。"
    return head + tail, note
