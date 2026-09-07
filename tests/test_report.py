"""HTML 报告测试：自包含、转义、必需内容与修正记录。"""

from __future__ import annotations

import os
import re
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from tradecheck.models import Correction
from tradecheck.pipeline import check_pair, extract_pair, render_report_file
from tradecheck.report import build_report

A = os.path.join(ROOT, "samples", "A")
def make_report(invoice_name="normal_invoice.xlsx", packing_name="normal_packing.xlsx",
                template_id="A", corrections=None):
    inv, pac, tpl = extract_pair(
        os.path.join(A, invoice_name), os.path.join(A, packing_name), template_id
    )
    res, extra = check_pair(inv, pac, tpl, corrections)
    html = build_report(inv, pac, res, tpl, extra, corrections)
    return html, inv, pac, res, tpl, extra


class TestReportContent(unittest.TestCase):
    def setUp(self):
        self.html, self.inv, self.pac, self.res, self.tpl, self.extra = make_report()

    def test_contains_sha256(self):
        self.assertIn(self.inv.sha256, self.html)
        self.assertIn(self.pac.sha256, self.html)

    def test_contains_timestamp_with_timezone(self):
        # 形如 2026-09-05T09:30:00+07:00（CST / UTC+0700）
        self.assertTrue(
            re.search(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[+-]\d{2}:\d{2}", self.html),
            "报告必须包含带时区的检查时间",
        )

    def test_contains_tool_version(self):
        from tradecheck.version import VERSION

        self.assertIn(VERSION, self.html)

    def test_contains_rule_config(self):
        self.assertIn("rounding", self.html)
        self.assertIn("tolerance", self.html)
        self.assertIn("half_up", self.html)

    def test_contains_coverage_and_unchecked_sections(self):
        self.assertIn("检查覆盖范围", self.html)
        self.assertIn("未检查 / 需人工处理项", self.html)

    def test_contains_cell_locations(self):
        self.assertIn("Invoice!E8", self.html)
        self.assertIn("Packing List!E8", self.html)

    def test_overall_is_hedged(self):
        self.assertIn("已检查范围", self.res.overall)
        self.assertIn("人工复核", self.res.overall_note)

    def test_self_contained_no_external_resources(self):
        self.assertNotIn("http://", self.html)
        self.assertNotIn("https://", self.html)
        self.assertNotIn("<script", self.html.lower())
        self.assertNotIn("cdn", self.html.lower().replace("cdnt", ""))


class TestReportEscaping(unittest.TestCase):
    def test_malicious_model_is_escaped(self):
        """型号里塞进脚本标签，必须被转义，不能原样进入 HTML。"""
        inv, pac, tpl = extract_pair(
            os.path.join(A, "normal_invoice.xlsx"), os.path.join(A, "normal_packing.xlsx"), "A"
        )
        inv.rows[0].fields["model"].text = "<script>alert('x')</script>"
        inv.rows[0].fields["model"].key = "<script>alert('x')</script>"
        res, extra = check_pair(inv, pac, tpl)
        html = build_report(inv, pac, res, tpl, extra)
        self.assertNotIn("<script>", html)
        self.assertIn("&lt;script&gt;", html)

    def test_ampersand_and_quotes_escaped(self):
        inv, pac, tpl = extract_pair(
            os.path.join(A, "normal_invoice.xlsx"), os.path.join(A, "normal_packing.xlsx"), "A"
        )
        inv.header["doc_no"].text = 'A & B "C" <D>'
        res, extra = check_pair(inv, pac, tpl)
        html = build_report(inv, pac, res, tpl, extra)
        self.assertNotIn('A & B "C" <D>', html)
        self.assertIn("A &amp; B", html)


class TestReportPendingVisibility(unittest.TestCase):
    def test_pending_items_get_a_warning_banner(self):
        _, _, _, res, _, _ = make_report(
            "unit_invoice.xlsx", "unit_packing.xlsx"
        )
        self.assertGreater(res.count("待确认"), 0)
        html, *_ = make_report("unit_invoice.xlsx", "unit_packing.xlsx")
        self.assertIn("不能", html)
        self.assertIn("待确认", html)

    def test_green_only_conclusion_is_not_possible_with_pending(self):
        html, inv, pac, res, tpl, extra = make_report("unit_invoice.xlsx", "unit_packing.xlsx")
        self.assertNotIn("全部通过", html)


class TestReportFile(unittest.TestCase):
    def test_write_report_file(self):
        inv, pac, tpl = extract_pair(
            os.path.join(A, "qty_short_invoice.xlsx"), os.path.join(A, "qty_short_packing.xlsx"), "A"
        )
        corr = [Correction(target="invoice:row.8.qty", old_text="1000", new_text="980")]
        res, extra = check_pair(inv, pac, tpl, corr)
        with tempfile.TemporaryDirectory() as temp_dir:
            path = render_report_file(
                os.path.join(temp_dir, "test-report.html"), inv, pac, res, tpl, extra, corr
            )
            self.assertTrue(os.path.isfile(path))
            with open(path, "r", encoding="utf-8") as fh:
                text = fh.read()
            self.assertIn("修正记录", text)
            self.assertIn("980", text)
            self.assertGreater(len(text), 5000)


if __name__ == "__main__":
    unittest.main(verbosity=2)
