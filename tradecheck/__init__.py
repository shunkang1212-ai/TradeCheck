"""TradeCheck · 外贸单据一致性检查器。

业务逻辑分层（均不依赖界面）：
  excelio  读取 .xlsx（双读识别公式缓存、SHA-256、大小/行数限制）
  extract  模板驱动提取关键字段与明细
  rules    R01–R06 一致性检查（金额用 Decimal）
  report   自包含 HTML 报告生成
  web      本地网页界面（仅 127.0.0.1）
"""

from .version import TOOL_NAME, TOOL_NAME_ZH, VERSION

__all__ = ["TOOL_NAME", "TOOL_NAME_ZH", "VERSION"]
