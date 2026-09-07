"""独立 HTML 报告生成。

约束（任务书 5. / 6.）：
- 单文件自包含，不引用任何外部 CDN、字体或脚本。
- 所有来自文件的内容一律 HTML 转义后再写入。
- 必须包含：输入文件 SHA-256、带时区的检查时间、工具版本、规则配置、
  检查覆盖范围、未检查项、修正记录。
"""

from __future__ import annotations

import html
import json
from datetime import datetime
from typing import Any, Dict, List, Optional

from .corrections import apply_corrections, clone_document
from .models import (
    STATUS_DIFF,
    STATUS_OK,
    STATUS_SKIP,
    STATUS_TODO,
    CheckResult,
    Correction,
    DocumentExtraction,
)
from .templates import money_config
from .version import TOOL_NAME, TOOL_NAME_ZH, VERSION

STATUS_COLORS = {
    STATUS_DIFF: "#b3261e",
    STATUS_TODO: "#8a5300",
    STATUS_SKIP: "#455a64",
    STATUS_OK: "#1b5e20",
}

STATUS_BG = {
    STATUS_DIFF: "#fce8e6",
    STATUS_TODO: "#fef3e2",
    STATUS_SKIP: "#eceff1",
    STATUS_OK: "#e6f4ea",
}


def e(value: Any) -> str:
    """HTML 转义。报告里所有外部数据都必须过这一层。"""
    return html.escape("" if value is None else str(value), quote=True)


def _now_text() -> str:
    now = datetime.now().astimezone()
    return f"{now.isoformat(timespec='seconds')}（{now.tzname()} / UTC{now.strftime('%z')}）"


def _badge(status: str) -> str:
    color = STATUS_COLORS.get(status, "#455a64")
    bg = STATUS_BG.get(status, "#eceff1")
    return (
        f'<span class="badge" style="color:{color};background:{bg};border-color:{color}">'
        f"{e(status)}</span>"
    )


def _doc_table(doc: DocumentExtraction) -> str:
    rows: List[str] = []
    for r in doc.rows:
        cells = []
        for name in ("model", "spec", "unit", "qty", "unit_price", "amount"):
            if name not in r.fields and not (doc.role == "invoice" and name in ("unit_price", "amount")):
                continue
            f = r.fields.get(name)
            coord = f.cell.coord if f.cell else "—"
            text = f.text if f.text else "—"
            mark = ' <span class="tag-formula">公式</span>' if f.is_formula else ""
            cells.append(f"<td>{e(text)}{mark}<div class='coord'>{e(coord)}</div></td>")
        rows.append(
            f"<tr><td class='rowno'>{int(r.row_index)}</td>" + "".join(cells) + "</tr>"
        )
    if not rows:
        return "<p class='muted'>（未提取到明细行）</p>"

    headers = ["行号", "型号", "规格", "单位", "数量"]
    if doc.role == "invoice":
        headers += ["单价", "行金额"]
    thead = "".join(f"<th>{e(h)}</th>" for h in headers)
    return (
        f"<table class='grid'><thead><tr>{thead}</tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table>"
    )


def _findings_table(result: CheckResult) -> str:
    order = {STATUS_DIFF: 0, STATUS_TODO: 1, STATUS_SKIP: 2, STATUS_OK: 3}
    findings = sorted(result.findings, key=lambda f: (order.get(f.status, 9), f.rule_id, f.item))
    rows: List[str] = []
    for f in findings:
        sources = "；".join(e(s.text()) for s in f.sources) if f.sources else "—"
        rows.append(
            "<tr>"
            f"<td class='nowrap'>{e(f.rule_id)}</td>"
            f"<td>{e(f.item)}</td>"
            f"<td class='nowrap'>{_badge(f.status)}</td>"
            f"<td>{e(f.expected) or '—'}</td>"
            f"<td>{e(f.actual) or '—'}</td>"
            f"<td class='evidence'>{e(f.evidence) or '—'}</td>"
            f"<td>{e(f.detail) or '—'}</td>"
            f"<td class='coord'>{sources}</td>"
            "</tr>"
        )
    if not rows:
        return "<p class='muted'>（无检查结果）</p>"
    thead = "".join(
        f"<th>{e(h)}</th>"
        for h in ["规则", "检查项目", "状态", "期望值", "实际值", "计算过程 / 依据", "差异说明", "位置（文件名 · 工作表!单元格）"]
    )
    return (
        f"<table class='grid findings'><thead><tr>{thead}</tr></thead><tbody>"
        + "".join(rows)
        + "</tbody></table>"
    )


def build_report(
    invoice: DocumentExtraction,
    packing: DocumentExtraction,
    result: CheckResult,
    template: Dict[str, Any],
    extra: Optional[Dict[str, Any]] = None,
    corrections: Optional[List[Correction]] = None,
) -> str:
    extra = extra or {}
    if corrections:
        # 报告展示的必须是修正后的数据；在副本上应用，不动传入对象
        invoice = clone_document(invoice)
        packing = clone_document(packing)
        apply_corrections({"invoice": invoice, "packing": packing}, corrections)
    money = money_config(template)
    summary = result.summary()

    warn = ""
    if summary[STATUS_TODO] or summary[STATUS_SKIP]:
        warn = (
            "<div class='callout warn'>"
            f"本结论<strong>不能</strong>视为“已完全一致、可放心放行”：存在 "
            f"{summary[STATUS_TODO]} 项待确认、{summary[STATUS_SKIP]} 项未检查。"
            "请先看下面的「待确认 / 未检查」条目，再决定是否放行这批单据。"
            "</div>"
        )
    if result.corrections:
        warn += (
            "<div class='callout'>"
            f"本报告基于用户修正后的数据（{len(result.corrections)} 处修正），"
            "结论不代表原始单据的自动提取结果。原文件未被修改。"
            "</div>"
        )

    corrections_rows = "".join(
        "<tr>"
        f"<td>{e(c.target)}</td>"
        f"<td>{e(c.old_text) or '（空）'}</td>"
        f"<td>{e(c.new_text) or '（空）'}</td>"
        f"<td class='coord'>{e(c.source)}</td>"
        "</tr>"
        for c in result.corrections
    )
    corrections_block = (
        "<h2>修正记录（仅本次会话，原文件未改动）</h2>"
        "<table class='grid'><thead><tr><th>修正位置</th><th>原值</th><th>新值</th><th>来源</th></tr></thead>"
        f"<tbody>{corrections_rows}</tbody></table>"
        if result.corrections
        else "<h2>修正记录</h2><p class='muted'>本次未对提取值做人工修正。</p>"
    )

    def bullets(items: List[str], empty: str) -> str:
        if not items:
            return f"<p class='muted'>{e(empty)}</p>"
        return "<ul>" + "".join(f"<li>{e(x)}</li>" for x in items) + "</ul>"

    rule_config = {
        "template_id": template.get("template_id"),
        "template_name": template.get("display_name"),
        "money": money,
        "unit_equivalences": {},
        "supported_checks": ["R01", "R02", "R03", "R04", "R05", "R06"],
        "not_supported": [
            "PDF / 图片 / Word 单据",
            "OCR 与任意布局自动识别",
            "重量、体积、毛净重核对",
            "复杂混装箱与分批出货",
            "折扣、税费与附加费用的自动计算",
            "单位换算（如 箱 -> 件）",
            "多币种混合计算",
        ],
    }

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(TOOL_NAME)} 检查报告 · {e(invoice.file_name)} / {e(packing.file_name)}</title>
<style>
  :root {{ color-scheme: light; }}
  body {{
    margin:0; padding:28px 32px 64px;
    font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Hiragino Sans GB","Microsoft YaHei",sans-serif;
    color:#1f2328; background:#ffffff; line-height:1.6;
  }}
  h1 {{ font-size:22px; margin:0 0 4px; }}
  h2 {{ font-size:17px; margin:32px 0 10px; padding-bottom:6px; border-bottom:1px solid #d8dee4; }}
  h3 {{ font-size:15px; margin:20px 0 8px; }}
  .sub {{ color:#57606a; font-size:13px; margin-bottom:18px; }}
  table.grid {{ border-collapse:collapse; width:100%; font-size:13px; margin:10px 0; }}
  table.grid th, table.grid td {{
    border:1px solid #d8dee4; padding:6px 9px; vertical-align:top; text-align:left;
  }}
  table.grid thead th {{ background:#f6f8fa; font-weight:600; white-space:nowrap; }}
  table.findings tbody tr:nth-child(even) {{ background:#fbfcfd; }}
  .badge {{
    display:inline-block; padding:1px 8px; border-radius:10px;
    border:1px solid; font-size:12px; white-space:nowrap; font-weight:600;
  }}
  .nowrap {{ white-space:nowrap; }}
  .coord {{ color:#57606a; font-size:11px; font-family:ui-monospace,SFMono-Regular,Menlo,monospace; word-break:break-all; }}
  .evidence {{ font-size:12px; color:#24292f; }}
  .rowno {{ color:#57606a; text-align:right; }}
  .tag-formula {{
    display:inline-block; font-size:10px; padding:0 4px; border:1px solid #8a5300;
    color:#8a5300; border-radius:3px; margin-left:4px;
  }}
  .overall {{
    border:1px solid #d8dee4; border-left:4px solid #1f6feb; background:#f6f8fa;
    padding:12px 16px; margin:16px 0; border-radius:4px;
  }}
  .overall .head {{ font-size:16px; font-weight:600; }}
  .overall .note {{ font-size:13px; color:#57606a; margin-top:6px; }}
  .callout {{
    border:1px solid #d8dee4; background:#f6f8fa; padding:10px 14px;
    border-radius:4px; margin:12px 0; font-size:13px;
  }}
  .callout.warn {{ border-color:#8a5300; background:#fef3e2; color:#6b4200; }}
  .stats span {{
    display:inline-block; margin-right:18px; font-size:13px; font-weight:600;
  }}
  .muted {{ color:#57606a; font-size:13px; }}
  code, pre {{
    font-family:ui-monospace,SFMono-Regular,Menlo,monospace; font-size:12px;
    background:#f6f8fa; border:1px solid #d8dee4; border-radius:4px;
  }}
  pre {{ padding:10px 12px; overflow-x:auto; white-space:pre-wrap; word-break:break-all; }}
  .kv td:first-child {{ width:190px; color:#57606a; background:#f6f8fa; }}
  footer {{ margin-top:36px; padding-top:12px; border-top:1px solid #d8dee4; color:#57606a; font-size:12px; }}
</style>
</head>
<body>

<h1>{e(TOOL_NAME)} · {e(TOOL_NAME_ZH)} — 检查报告</h1>
<div class="sub">
  工具版本 {e(VERSION)} · 模板 {e(template.get('template_id'))}（{e(template.get('display_name'))}） ·
  检查时间 {e(_now_text())}
</div>

<div class="overall">
  <div class="head">{e(result.overall)}</div>
  <div class="stats" style="margin-top:8px">
    <span style="color:{STATUS_COLORS[STATUS_DIFF]}">发现差异 {summary[STATUS_DIFF]}</span>
    <span style="color:{STATUS_COLORS[STATUS_TODO]}">待确认 {summary[STATUS_TODO]}</span>
    <span style="color:{STATUS_COLORS[STATUS_SKIP]}">未检查 {summary[STATUS_SKIP]}</span>
    <span style="color:{STATUS_COLORS[STATUS_OK]}">未发现差异 {summary[STATUS_OK]}</span>
  </div>
  <div class="note">{e(result.overall_note)}</div>
</div>
{warn}

<h2>1. 输入文件</h2>
<table class="grid kv">
  <tbody>
    <tr><td>商业发票文件</td><td>{e(invoice.file_name)}</td></tr>
    <tr><td>商业发票 SHA-256</td><td class='coord'>{e(invoice.sha256)}</td></tr>
    <tr><td>商业发票工作表</td><td>{e(invoice.sheet)}</td></tr>
    <tr><td>装箱单文件</td><td>{e(packing.file_name)}</td></tr>
    <tr><td>装箱单 SHA-256</td><td class='coord'>{e(packing.sha256)}</td></tr>
    <tr><td>装箱单工作表</td><td>{e(packing.sheet)}</td></tr>
    <tr><td>币种识别结果</td><td>{e(extra.get('currency_text') or '—')}（{e(extra.get('currency_message') or '—')}）</td></tr>
  </tbody>
</table>
<p class="muted">SHA-256 于读取前计算，用于核对报告对应的文件版本；本工具不会修改输入文件。</p>

<h2>2. 检查结果</h2>
{_findings_table(result)}

<h2>3. 检查覆盖范围</h2>
{bullets(result.coverage, '本次没有可覆盖的检查项（输入可能未能正确读取）。')}

<h2>4. 未检查 / 需人工处理项</h2>
{bullets(result.unchecked, '无。')}

{corrections_block}

<h2>5. 提取到的关键字段</h2>
<table class="grid">
  <thead><tr><th>字段</th><th>商业发票</th><th>位置</th><th>装箱单</th><th>位置</th></tr></thead>
  <tbody>
    <tr>
      <td>单据号</td>
      <td>{e(invoice.header_field('doc_no').text) or '—'}</td>
      <td class='coord'>{e(invoice.header_field('doc_no').cell.coord) if invoice.header_field('doc_no').cell else '—'}</td>
      <td>{e(packing.header_field('invoice_ref').text) or '—'}</td>
      <td class='coord'>{e(packing.header_field('invoice_ref').cell.coord) if packing.header_field('invoice_ref').cell else '—'}</td>
    </tr>
    <tr>
      <td>订单引用号</td>
      <td>{e(invoice.header_field('order_ref').text) or '—'}</td>
      <td class='coord'>{e(invoice.header_field('order_ref').cell.coord) if invoice.header_field('order_ref').cell else '—'}</td>
      <td>{e(packing.header_field('order_ref').text) or '—'}</td>
      <td class='coord'>{e(packing.header_field('order_ref').cell.coord) if packing.header_field('order_ref').cell else '—'}</td>
    </tr>
    <tr>
      <td>币种</td>
      <td>{e(invoice.header_field('currency').text) or '—'}</td>
      <td class='coord'>{e(invoice.header_field('currency').cell.coord) if invoice.header_field('currency').cell else '—'}</td>
      <td class='muted'>（装箱单不参与金额）</td><td>—</td>
    </tr>
    <tr>
      <td>单据合计</td>
      <td>{e(invoice.total.text) if invoice.total else '—'}</td>
      <td class='coord'>{e(invoice.total.cell.coord) if invoice.total and invoice.total.cell else '—'}</td>
      <td>{e(packing.total.text) if packing.total else '—'}</td>
      <td class='coord'>{e(packing.total.cell.coord) if packing.total and packing.total.cell else '—'}</td>
    </tr>
  </tbody>
</table>

<h2>6. 提取到的明细（原文保留，含单元格位置）</h2>
<h3>商业发票 · {e(invoice.sheet)}</h3>
{_doc_table(invoice)}
<h3>装箱单 · {e(packing.sheet)}</h3>
{_doc_table(packing)}

<h2>7. 规则配置</h2>
<pre>{e(json.dumps(rule_config, ensure_ascii=False, indent=2))}</pre>
<p class="muted">
  金额规则：币种 {e(money.get('default_currency'))}，保留 {e(money.get('digits'))} 位小数，
  舍入 {e(money.get('rounding_mode'))}，容差 {e(money.get('tolerance_absolute'))}。
  这是演示模板的约定，不代表所有币种的通用规则；实际小数位数按识别到的币种确定，
  未知币种一律降级为待确认，不假定两位小数。
</p>

<h2>8. 限制与免责</h2>
<ul>
  <li>本报告只覆盖「已检查」范围内的项目，不等同于全量审单，也不代表 100% 可靠。</li>
  <li>存在「待确认 / 未检查」条目时，不得据此判定单据无误。</li>
  <li>重要金额、币种及出货数据仍需用户人工复核。</li>
  <li>本工具不修改原始单据，不联网，不上传文件内容。</li>
  <li>测试用例基于虚构样例，不作为真实业务准确率依据。</li>
</ul>

<footer>
  {e(TOOL_NAME)} {e(VERSION)} · 报告自包含，不依赖外部资源 · 生成时间 {e(_now_text())}
</footer>
</body>
</html>
"""


def write_report(path: str, html_text: str) -> str:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html_text)
    return path
