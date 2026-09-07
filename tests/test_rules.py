"""规则与流程的自动化测试，对应任务书第 7 节的 12 项验收点。

预期结果全部独立硬编码，不调用被测函数生成（任务书 7.）。
"""

from __future__ import annotations

import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from tradecheck.excelio import sha256_file
from tradecheck.models import STATUS_DIFF, STATUS_OK, STATUS_SKIP, STATUS_TODO, Correction
from tradecheck.pipeline import check_pair, extract_pair
from tradecheck.report import build_report
from tradecheck.templates import load_template

SAMPLES = os.path.join(ROOT, "samples")
A = os.path.join(SAMPLES, "A")
B = os.path.join(SAMPLES, "B")


def load(name_invoice: str, name_packing: str, template_id: str):
    return extract_pair(
        os.path.join(A if template_id == "A" else B, name_invoice),
        os.path.join(A if template_id == "A" else B, name_packing),
        template_id,
    )


class Base(unittest.TestCase):
    def statuses(self, result, rule_id=None):
        return [
            f.status for f in result.findings if rule_id is None or f.rule_id == rule_id
        ]

    def items(self, result, rule_id=None):
        return [f.item for f in result.findings if rule_id is None or f.rule_id == rule_id]


# ---------------------------------------------------------------- 1 正常单据


class Test01Normal(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("normal_invoice.xlsx", "normal_packing.xlsx", "A")
        self.res, self.extra = check_pair(self.inv, self.pac, self.tpl)

    def test_no_false_positive(self):
        """正常单据不得误报：既不能有差异，也不应留下待确认/未检查。"""
        self.assertEqual(self.res.count(STATUS_DIFF), 0)
        self.assertEqual(self.res.count(STATUS_TODO), 0)
        self.assertEqual(self.res.count(STATUS_SKIP), 0)
        self.assertIn("在已检查范围内未发现差异", self.res.overall)

    def test_all_three_checks_ran(self):
        self.assertEqual(len(self.statuses(self.res, "R03")), 3)  # 3 个产品
        self.assertEqual(len(self.statuses(self.res, "R04")), 3)  # 3 行金额
        self.assertEqual(len(self.statuses(self.res, "R05")), 1)  # 总额

    def test_extracted_values(self):
        self.assertEqual(self.inv.header_field("doc_no").text, "TCI-2026-0001")
        self.assertEqual(self.inv.header_field("currency").text, "USD")
        self.assertEqual(self.inv.total.text, "3150")
        self.assertEqual(self.inv.sheet, "Invoice")
        self.assertEqual(self.pac.sheet, "Packing List")

    def test_currency_resolved(self):
        self.assertEqual(self.extra["currency_code"], "USD")
        self.assertEqual(self.extra["currency_digits"], 2)


# ---------------------------------------------------------------- 2 短装 20 件


class Test02QuantityShort(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("qty_short_invoice.xlsx", "qty_short_packing.xlsx", "A")
        self.res, _ = check_pair(self.inv, self.pac, self.tpl)

    def test_difference_reported(self):
        diffs = [f for f in self.res.findings if f.status == STATUS_DIFF]
        self.assertEqual(len(diffs), 1)
        f = diffs[0]
        self.assertEqual(f.rule_id, "R03")

    def test_difference_value_is_20(self):
        f = [x for x in self.res.findings if x.status == STATUS_DIFF][0]
        self.assertIn("-20", f.evidence)
        self.assertIn("1,000", f.expected)
        self.assertIn("980", f.actual)

    def test_source_cells_are_correct(self):
        """发票数量在 Invoice!E8，装箱单数量在 Packing List!E8。"""
        f = [x for x in self.res.findings if x.status == STATUS_DIFF][0]
        src = f.source_text()
        self.assertIn("Invoice!E8", src)
        self.assertIn("Packing List!E8", src)

    def test_not_silently_passed(self):
        self.assertEqual(self.res.count(STATUS_DIFF), 1)


# ---------------------------------------------------------------- 3 拆行汇总


class Test03SplitRows(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("split_invoice.xlsx", "split_packing.xlsx", "A")
        self.res, _ = check_pair(self.inv, self.pac, self.tpl)

    def test_packing_has_more_rows_than_invoice(self):
        self.assertEqual(len(self.inv.rows), 1)
        self.assertEqual(len(self.pac.rows), 2)

    def test_summed_then_matched(self):
        self.assertEqual(self.res.count(STATUS_DIFF), 0)
        f = [x for x in self.res.findings if x.rule_id == "R03" and "TC-100" in x.item][0]
        self.assertEqual(f.status, STATUS_OK)
        self.assertIn("600", f.evidence)
        self.assertIn("400", f.evidence)

    def test_all_source_cells_listed(self):
        """汇总值必须列出所有参与汇总的来源单元格。"""
        f = [x for x in self.res.findings if x.rule_id == "R03" and "TC-100" in x.item][0]
        src = f.source_text()
        self.assertIn("Packing List!E8", src)
        self.assertIn("Packing List!E9", src)


# ---------------------------------------------------------------- 4 同型号不同规格


class Test04SameModelDifferentSpec(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("spec_invoice.xlsx", "spec_packing.xlsx", "A")
        self.res, _ = check_pair(self.inv, self.pac, self.tpl)

    def test_not_merged(self):
        """若错误合并成 TC-100 = 800 vs 800，会漏报；这里必须是 2 项差异。"""
        self.assertEqual(self.res.count(STATUS_DIFF), 2)

    def test_both_specs_reported_separately(self):
        items = " ".join(f.item for f in self.res.findings if f.status == STATUS_DIFF)
        self.assertIn("RED-L", items)
        self.assertIn("BLU-M", items)

    def test_differences_have_opposite_signs(self):
        """发票 RED 500 / BLU 300，装箱单 RED 300 / BLU 500。"""
        ev = " ".join(f.evidence for f in self.res.findings if f.status == STATUS_DIFF)
        self.assertIn("200", ev)
        self.assertIn("-200", ev)


# ---------------------------------------------------------------- 5 顺序不同


class Test05RowOrder(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("reorder_invoice.xlsx", "reorder_packing.xlsx", "A")
        self.res, _ = check_pair(self.inv, self.pac, self.tpl)

    def test_matched_regardless_of_order(self):
        self.assertEqual(self.res.count(STATUS_DIFF), 0)
        self.assertEqual(self.res.count(STATUS_TODO), 0)
        self.assertEqual(self.res.count(STATUS_SKIP), 0)

    def test_all_products_compared(self):
        r03 = [f for f in self.res.findings if f.rule_id == "R03"]
        self.assertEqual(len(r03), 3)
        self.assertTrue(all(f.status == STATUS_OK for f in r03))


# ---------------------------------------------------------------- 6 件 vs 箱


class Test06UnitMismatch(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("unit_invoice.xlsx", "unit_packing.xlsx", "A")
        self.res, _ = check_pair(self.inv, self.pac, self.tpl)

    def test_unit_difference_is_flagged(self):
        todos = [f for f in self.res.findings if f.status == STATUS_TODO]
        self.assertEqual(len(todos), 1)
        self.assertEqual(todos[0].rule_id, "R03")
        self.assertIn("单位", todos[0].item)

    def test_quantity_not_compared(self):
        """1000 与 1000 数值相同，但单位不同，不得据此判定一致。"""
        self.assertEqual(self.res.count(STATUS_DIFF), 0)
        self.assertNotIn("数量比较：TC-700 / STD", [f.item for f in self.res.findings])

    def test_no_conversion_ratio_guessed(self):
        f = [x for x in self.res.findings if x.status == STATUS_TODO][0]
        self.assertIn("换算", f.detail)

    def test_recorded_as_unchecked(self):
        self.assertTrue(any("单位不一致" in u for u in self.res.unchecked))


# ---------------------------------------------------------------- 7 行金额错误


class Test07LineAmount(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("amount_invoice.xlsx", "amount_packing.xlsx", "A")
        self.res, _ = check_pair(self.inv, self.pac, self.tpl)

    def test_error_row_reported(self):
        diffs = [f for f in self.res.findings if f.status == STATUS_DIFF]
        self.assertEqual(len(diffs), 1)
        self.assertEqual(diffs[0].rule_id, "R04")

    def test_computation_is_reproducible(self):
        """120 × 12.50 = 1500.00，单据写 1250.00，差 -250.00。"""
        f = [x for x in self.res.findings if x.status == STATUS_DIFF][0]
        self.assertIn("1,500.00", f.expected)
        self.assertIn("1,250.00", f.actual)
        self.assertIn("120 × 12.50", f.evidence)
        self.assertIn("-250.00", f.evidence)

    def test_correct_row_still_passes(self):
        ok = [f for f in self.res.findings if f.rule_id == "R04" and f.status == STATUS_OK]
        self.assertEqual(len(ok), 1)

    def test_total_still_consistent(self):
        """总额按单据行金额加起来仍等于单据总额，不应误报。"""
        r05 = [f for f in self.res.findings if f.rule_id == "R05"]
        self.assertEqual(len(r05), 1)
        self.assertEqual(r05[0].status, STATUS_OK)


# ---------------------------------------------------------------- 8 舍入边界


class Test08Rounding(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("round_invoice.xlsx", "round_packing.xlsx", "A")
        self.res, _ = check_pair(self.inv, self.pac, self.tpl)

    def test_half_up_and_decimal_precision(self):
        """0.125->0.13（HALF_UP）、2.675->2.68（Decimal 精确，float 会算成 2.67）。"""
        self.assertEqual(self.res.count(STATUS_DIFF), 0)
        self.assertEqual(self.res.count(STATUS_TODO), 0)
        self.assertEqual(self.res.count(STATUS_SKIP), 0)

    def test_repeatable_across_runs(self):
        outs = []
        for _ in range(5):
            inv, pac, tpl = load("round_invoice.xlsx", "round_packing.xlsx", "A")
            res, _ = check_pair(inv, pac, tpl)
            outs.append((res.overall, tuple(sorted(f.item + f.status for f in res.findings))))
        self.assertEqual(len(set(outs)), 1, "同样输入多次运行结果必须一致")

    def test_total_matches(self):
        r05 = [f for f in self.res.findings if f.rule_id == "R05"][0]
        self.assertEqual(r05.status, STATUS_OK)
        self.assertIn("3.11", r05.expected)


# ---------------------------------------------------------------- 9 缺字段


class Test09MissingFields(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("missing_invoice.xlsx", "missing_packing.xlsx", "A")
        self.res, self.extra = check_pair(self.inv, self.pac, self.tpl)

    def test_missing_currency_blocks_amount_checks(self):
        self.assertIsNone(self.extra["currency_code"])
        r04 = [f for f in self.res.findings if f.rule_id == "R04"]
        self.assertTrue(r04, "缺少币种必须有明确提示")
        self.assertEqual(r04[0].status, STATUS_TODO)
        self.assertFalse(
            any(f.rule_id == "R04" and f.status == STATUS_OK for f in self.res.findings),
            "币种未知时不得判定通过",
        )

    def test_total_check_skipped(self):
        r05 = [f for f in self.res.findings if f.rule_id == "R05"]
        self.assertEqual(r05[0].status, STATUS_SKIP)

    def test_missing_model_flagged(self):
        self.assertTrue(any("型号" in f.item and f.status == STATUS_TODO for f in self.res.findings))

    def test_non_numeric_qty_flagged(self):
        self.assertTrue(
            any("N/A" in f.actual or "N/A" in f.evidence for f in self.res.findings)
        )

    def test_quantity_compare_skipped_for_bad_row(self):
        self.assertTrue(any("TC-800" in f.item and f.status == STATUS_SKIP for f in self.res.findings))

    def test_overall_mentions_pending_items(self):
        self.assertIn("待确认", self.res.overall)


# ---------------------------------------------------------------- 10 公式缓存


class Test10FormulaCache(Base):
    def test_missing_cache_is_not_a_reliable_value(self):
        inv, pac, tpl = load("formula_nocache_invoice.xlsx", "formula_nocache_packing.xlsx", "A")
        res, _ = check_pair(inv, pac, tpl)
        r04 = [f for f in res.findings if f.rule_id == "R04"]
        self.assertEqual(len(r04), 2)
        self.assertTrue(all(f.status == STATUS_SKIP for f in r04))
        self.assertTrue(all("公式" in f.evidence for f in r04))

    def test_stale_cache_is_not_trusted(self):
        """缓存被人为改成 9999/8888/7777，正确行为是「未检查」，不是「发现差异」。"""
        inv, pac, tpl = load("formula_stale_invoice.xlsx", "formula_stale_packing.xlsx", "A")
        res, _ = check_pair(inv, pac, tpl)
        r04 = [f for f in res.findings if f.rule_id == "R04"]
        self.assertEqual(len(r04), 2)
        self.assertTrue(all(f.status == STATUS_SKIP for f in r04))
        self.assertEqual(res.count(STATUS_DIFF), 0, "过期缓存不得当成可靠差异依据")

    def test_formula_source_preserved(self):
        inv, pac, tpl = load("formula_stale_invoice.xlsx", "formula_stale_packing.xlsx", "A")
        res, _ = check_pair(inv, pac, tpl)
        r06 = [f for f in res.findings if f.rule_id == "R06"]
        self.assertTrue(any("=E8*F8" in (f.actual or "") for f in r06))

    def test_quantity_still_checked_when_constants(self):
        """装箱单是常量，发票数量也是常量，数量比较应正常执行。"""
        inv, pac, tpl = load("formula_stale_invoice.xlsx", "formula_stale_packing.xlsx", "A")
        res, _ = check_pair(inv, pac, tpl)
        r03 = [f for f in res.findings if f.rule_id == "R03"]
        self.assertTrue(all(f.status == STATUS_OK for f in r03))

    def test_cells_are_marked_unverified(self):
        inv, pac, tpl = load("formula_stale_invoice.xlsx", "formula_stale_packing.xlsx", "A")
        amount = inv.rows[0].get("amount")
        self.assertTrue(amount.is_formula)
        self.assertEqual(amount.formula, "=E8*F8")
        self.assertIsNone(amount.usable_number, "公式缓存不得作为可计算数值")


# ---------------------------------------------------------------- 11 额外费用 / 未知模板


class Test11OutOfScope(Base):
    def test_extra_cost_structure_flagged(self):
        inv, pac, tpl = load("extracost_invoice.xlsx", "extracost_packing.xlsx", "A")
        res, _ = check_pair(inv, pac, tpl)
        f = [x for x in res.findings if "金额结构" in x.item]
        self.assertEqual(len(f), 1)
        self.assertEqual(f[0].status, STATUS_TODO)
        self.assertIn("Freight", f[0].evidence)
        self.assertIn("人工核对", f[0].detail)

    def test_extra_cost_total_is_not_passed(self):
        inv, pac, tpl = load("extracost_invoice.xlsx", "extracost_packing.xlsx", "A")
        res, _ = check_pair(inv, pac, tpl)
        r05 = [x for x in res.findings if x.rule_id == "R05" and x.item == "发票总额"]
        self.assertEqual(r05[0].status, STATUS_TODO)
        self.assertNotEqual(r05[0].status, STATUS_OK)

    def test_extra_cost_row_not_treated_as_product(self):
        """Freight 行不能被当成产品去比对数量。"""
        inv, pac, tpl = load("extracost_invoice.xlsx", "extracost_packing.xlsx", "A")
        self.assertEqual(len(inv.rows), 1)

    def test_unknown_currency(self):
        inv, pac, tpl = load("unknownccy_invoice.xlsx", "unknownccy_packing.xlsx", "A")
        res, extra = check_pair(inv, pac, tpl)
        self.assertIsNone(extra["currency_code"])
        self.assertEqual(res.count(STATUS_DIFF), 0, "币种未知时不应给出金额差异结论")

    def test_unsupported_template(self):
        path = os.path.join(SAMPLES, "unsupported_layout.xlsx")
        inv, pac, tpl = extract_pair(path, path, "A")
        res, _ = check_pair(inv, pac, tpl)
        self.assertEqual(inv.table_found, False)
        self.assertTrue(any(f.status == STATUS_SKIP for f in res.findings))
        joined = " ".join(f.actual + f.evidence for f in res.findings)
        self.assertIn("不支持", joined + " ".join(f.detail for f in res.findings))

    def test_wrong_template_selected(self):
        """中文单据（模板 B 布局）误选英文模板 A，必须明确提示，不能默默给出结果。"""
        inv, pac, tpl = extract_pair(
            os.path.join(B, "normal_invoice.xlsx"),
            os.path.join(B, "normal_packing.xlsx"),
            "A",
        )
        self.assertFalse(inv.table_found)
        res, _ = check_pair(inv, pac, tpl)
        self.assertTrue(
            any("模板" in f.actual or "模板" in f.detail for f in res.findings)
        )


# ---------------------------------------------------------------- 12 用户修正


class Test12Corrections(Base):
    def setUp(self):
        self.inv, self.pac, self.tpl = load("qty_short_invoice.xlsx", "qty_short_packing.xlsx", "A")

    def test_correction_fixes_the_difference(self):
        """把发票数量 1000 修正为 980 后，原来的数量差异（R03）应消失。

        注意：修正是“外科手术式”的——只改写被指定的那个提取值，
        不会自动同步其它独立字段。发票行金额是常量 2,500（=1000×2.5），
        按修正后的数量重算应为 2,450，因此 R04 可能浮现新的金额提示。
        这里只验证被修正的目标差异已消除，符合“原来的一项差异应消失”。
        """
        before, _ = check_pair(self.inv, self.pac, self.tpl)
        r03_before = [f for f in before.findings if f.rule_id == "R03"]
        self.assertTrue(
            any(f.status == STATUS_DIFF for f in r03_before),
            "修正前应存在 R03 数量差异",
        )

        corr = [Correction(target="invoice:row.8.qty", old_text="1000", new_text="980")]
        after, _ = check_pair(self.inv, self.pac, self.tpl, corr)
        r03_after = [f for f in after.findings if f.rule_id == "R03"]
        self.assertFalse(
            any(f.status == STATUS_DIFF for f in r03_after),
            "修正数量后，R03 数量差异应消失",
        )
        self.assertEqual(len(after.corrections), 1)
        self.assertEqual(after.corrections[0].old_text, "1000")
        self.assertEqual(after.corrections[0].new_text, "980")

    def test_correction_records_source(self):
        corr = [Correction(target="invoice:row.8.qty", old_text="1000", new_text="980")]
        after, _ = check_pair(self.inv, self.pac, self.tpl, corr)
        self.assertIn("Invoice!E8", after.corrections[0].source)

    def test_correction_does_not_mutate_original_extraction(self):
        corr = [Correction(target="invoice:row.8.qty", old_text="1000", new_text="980")]
        check_pair(self.inv, self.pac, self.tpl, corr)
        self.assertEqual(self.inv.rows[0].get("qty").text, "1000", "原提取结果不得被修改")

    def test_report_contains_correction_section(self):
        corr = [Correction(target="invoice:row.8.qty", old_text="1000", new_text="980")]
        res, extra = check_pair(self.inv, self.pac, self.tpl, corr)
        html = build_report(self.inv, self.pac, res, self.tpl, extra, corr)
        self.assertIn("修正记录", html)
        self.assertIn("980", html)
        self.assertIn("基于用户修正后的数据", html)

    def test_correction_to_bad_value_is_flagged(self):
        """把数量改成非数字，应变成待确认，而不是通过。"""
        corr = [Correction(target="invoice:row.8.qty", old_text="1000", new_text="待定")]
        res, _ = check_pair(self.inv, self.pac, self.tpl, corr)
        r03 = [f for f in res.findings if f.rule_id == "R03"]
        self.assertTrue(all(f.status != STATUS_OK for f in r03))


# ---------------------------------------------------------------- 前导零 / 模板 B / 文件保护


class TestModelText(Base):
    def test_leading_zero_preserved(self):
        inv, pac, tpl = load("zero_ok_invoice.xlsx", "zero_ok_packing.xlsx", "A")
        self.assertEqual(inv.rows[0].get("model").text, "007-A")
        self.assertEqual(pac.rows[0].get("model").text, "007-A")
        res, _ = check_pair(inv, pac, tpl)
        self.assertEqual(res.count(STATUS_DIFF), 0)

    def test_leading_zero_is_significant(self):
        """007-A 与 7-A 不是同一个产品，不得因数值化而合并。"""
        inv, pac, tpl = load("zero_mismatch_invoice.xlsx", "zero_mismatch_packing.xlsx", "A")
        res, _ = check_pair(inv, pac, tpl)
        self.assertEqual(res.count(STATUS_DIFF), 2)


class TestTemplateB(Base):
    def test_normal(self):
        inv, pac, tpl = load("normal_invoice.xlsx", "normal_packing.xlsx", "B")
        self.assertEqual(inv.sheet, "商业发票")
        self.assertEqual(pac.sheet, "装箱单")
        self.assertEqual(inv.header_field("currency").text, "CNY")
        res, extra = check_pair(inv, pac, tpl)
        self.assertEqual(extra["currency_code"], "CNY")
        self.assertEqual(res.count(STATUS_DIFF), 0)
        self.assertEqual(res.count(STATUS_TODO), 0)
        self.assertEqual(res.count(STATUS_SKIP), 0)

    def test_different_layout_positions(self):
        """模板 B 装箱单的列顺序与行起点都与发票不同。"""
        inv, pac, tpl = load("normal_invoice.xlsx", "normal_packing.xlsx", "B")
        self.assertEqual(inv.rows[0].row_index, 8)
        self.assertEqual(pac.rows[0].row_index, 7)
        self.assertEqual(inv.rows[0].get("qty").cell.coord, "E8")
        self.assertEqual(pac.rows[0].get("qty").cell.coord, "F7")

    def test_split_rows(self):
        inv, pac, tpl = load("split_invoice.xlsx", "split_packing.xlsx", "B")
        self.assertEqual(len(pac.rows), 3)
        res, _ = check_pair(inv, pac, tpl)
        self.assertEqual(res.count(STATUS_DIFF), 0)

    def test_quantity_short(self):
        inv, pac, tpl = load("qty_short_invoice.xlsx", "qty_short_packing.xlsx", "B")
        res, _ = check_pair(inv, pac, tpl)
        self.assertEqual(res.count(STATUS_DIFF), 1)
        f = [x for x in res.findings if x.status == STATUS_DIFF][0]
        self.assertIn("-50", f.evidence)


class TestFileIntegrity(Base):
    def test_input_files_not_modified_by_checking(self):
        paths = [
            os.path.join(A, "normal_invoice.xlsx"),
            os.path.join(A, "normal_packing.xlsx"),
            os.path.join(A, "formula_stale_invoice.xlsx"),
            os.path.join(B, "normal_invoice.xlsx"),
        ]
        before = {p: sha256_file(p) for p in paths}
        load("normal_invoice.xlsx", "normal_packing.xlsx", "A")
        load("formula_stale_invoice.xlsx", "formula_stale_packing.xlsx", "A")
        load("normal_invoice.xlsx", "normal_packing.xlsx", "B")
        after = {p: sha256_file(p) for p in paths}
        self.assertEqual(before, after, "检查过程不得修改输入文件")


if __name__ == "__main__":
    unittest.main(verbosity=2)
