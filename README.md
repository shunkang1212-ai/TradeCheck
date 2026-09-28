# TradeCheck

TradeCheck 是一个本地优先的外贸单据一致性检查器。它读取商业发票和装箱单的
`.xlsx` 文件，检查引用关系、型号/规格、数量、行金额和总额，并生成可留档的
单文件 HTML 报告。

项目当前为 **v0.1.1 alpha**。它适合测试和辅助复核，不能替代人工确认、财务审核、
报关判断或合同履行判断。

## 为什么做这个项目

中小外贸团队经常需要人工对照商业发票与装箱单。TradeCheck 把重复检查变成一组
明确规则，并保留文件哈希、工作表、单元格、修正记录和降级原因，方便复查。

- 本地处理：服务只监听 `127.0.0.1`，页面无外部 CDN。
- 原件不回写：用户修正只作用于当前会话的内存副本。
- 诚实降级：字段缺失、公式缓存、币种或单位不明确时标为“待确认/未检查”。
- 金额精度：使用 `Decimal` 和明确的舍入、容差规则。
- 可复现：仓库只提供脚本生成的虚构样例。

## 快速开始

需要 Python 3.9 或更高版本。

```bash
git clone https://github.com/shunkang1212-ai/TradeCheck.git
cd TradeCheck
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e .
tradecheck
```

Windows PowerShell 激活虚拟环境时使用：

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -e .
tradecheck
```

浏览器打开终端显示的 `http://127.0.0.1:端口/`。也可以指定端口：

```bash
tradecheck --port 9000
```

未安装项目时，也可以在仓库根目录运行 `python run.py`。

### 批量核对并输出 JSON

这是尚未发布的开发版候选功能；公开的 v0.1.1 Release 不包含批量命令。

批量模式读取 JSON 清单，每个条目显式配对一份商业发票和一份装箱单。相对文件路径按清单
所在目录解析；每组可以单独指定模板，省略时使用清单中的默认模板（默认 A）。仓库自带
一个使用虚构样例的清单：

```bash
tradecheck batch --manifest samples/batch-manifest.example.json --output results.json
```

也可以按此文件中的结构创建自己的清单。运行后会写出结构化结果：

```bash
tradecheck batch --manifest batch.json --output results.json
```

省略 `--output` 时 JSON 会写到标准输出，便于接入其他本地脚本。结果包含每组结论、规则
汇总、证据位置、文件名和 SHA-256；其中的期望值、实际值可能包含单据业务数据，请勿未经
检查就公开分享。结果文件若已存在，命令会拒绝覆盖；确认要替换结果文件时加 `--force`。
该保护也会阻止把 JSON 输出写到清单或输入工作簿路径。

退出码：`0` 表示所有组均为“未发现差异”；`1` 表示至少一组存在差异、待确认、未检查或
输入错误；`2` 表示清单、参数或输出路径无法使用。

## 使用流程

1. 选择模板 A（英文表头）或模板 B（中文表头）。
2. 上传商业发票和装箱单两份 `.xlsx` 文件。
3. 检查提取预览，必要时修正本次会话中的字段。
4. 确认两份文件属于同一批出货，再运行一致性核对。
5. 查看结论并生成本地 HTML 报告。

可以用 `samples/A/normal_invoice.xlsx` 与
`samples/A/normal_packing.xlsx` 完成第一次试用。样例中的型号、订单号和金额均为虚构。

## 检查规则

| 编号 | 检查 | 主要行为 |
| --- | --- | --- |
| R01 | 引用关联 | 比较发票号、订单号和相互引用 |
| R02 | 行匹配 | 按型号与规格配对，防止同型号不同规格误合并 |
| R03 | 数量 | 比较逐项数量；单位不一致时降级 |
| R04 | 行金额 | 用数量 × 单价重算，并按币种精度和容差比较 |
| R05 | 总额 | 比较发票总额；额外费用结构提示人工确认 |
| R06 | 输入完整性 | 缺失、无效、未验证公式影响的检查不会假装通过 |

结论优先级为：`发现差异` → `待确认` → `未检查` → `未发现差异`。
“未发现差异”只覆盖实际完成的检查项。

## 测试

```bash
python -m unittest discover -s tests
```

v0.1.1 在 Python 3.13 上通过 106 个自动化测试，并完成一次本机浏览器全流程检查。
GitHub Actions 会在 Python 3.9、3.12 和 3.13 上重新运行测试；跨平台结果以公开 CI
记录为准。这些验证对应 v0.1.1；尚未发布的批量 CLI 候选功能尚未经过自动化测试。

## 当前限制

- 只支持仓库内定义的两种 Excel 布局。
- 不支持 PDF、OCR、旧版 `.xls`、小数逗号、多币种和单位换算。
- 不重新计算 Excel 公式；依赖公式的检查会降级。
- 尚未使用真实业务样本验证，也未完成接近 20 MB/5000 行上限的性能测试。

请勿在公开 Issue 或 PR 中提交真实客户资料、订单、价格、银行信息、签章或联系人。
优先用 `tools/make_samples.py` 生成虚构的最小复现文件。

## 参与试用

项目正在收集首轮安装和可用性反馈。建议先用仓库中的虚构样例试用，并按
[试用指南](docs/PILOT_TEST_GUIDE.md)记录系统环境、操作结果和遇到的问题。请不要把真实
单据、生成的报告或包含客户信息的截图发到公开 Issue、PR 或其他公开渠道。

## 参与贡献

开始前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md) 和 [SECURITY.md](SECURITY.md)。
近期方向见 [ROADMAP.md](ROADMAP.md)。

## 许可证

[MIT](LICENSE)
