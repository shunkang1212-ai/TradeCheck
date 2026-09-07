#!/usr/bin/env python3
"""TradeCheck 启动入口。

用法：
    python run.py            # 默认 127.0.0.1:8753
    python run.py 9000       # 指定端口

启动后浏览器打开终端里给出的地址即可。仅本机可访问。
"""

import os
import sys

# 让 import 能找到 tradecheck 包（无论从哪里启动）
ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tradecheck.web import server  # noqa: E402


def main(argv):
    port = 8753
    if len(argv) > 1:
        try:
            port = int(argv[1])
        except ValueError:
            print(f"端口参数无效：{argv[1]}，使用默认 {port}")
    try:
        httpd = server.run(port=port)
    except RuntimeError as exc:
        print(str(exc))
        return 1
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        server._clear_session()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
